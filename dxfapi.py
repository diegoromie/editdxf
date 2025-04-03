import ezdxf
import os
import ezdxf.revcloud
import pandas as pd
import logging
from collections import defaultdict
from typing import Any, List
import re

# Set up logging configuration
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


# ===============================
#          LAYERS
# ===============================

def list_current_layers(doc: Any) -> List[str]:
    """
    Lists all layers present in the DXF document.

    :param doc: DXF document.
    :return: List of layer names.
    """
    return [layer.dxf.name for layer in doc.layers]


def create_layer(doc: Any, dest_layer: str, dest_color: int) -> None:
    """
    Creates a new layer in the DXF document if it does not already exist.

    :param doc: DXF document.
    :param dest_layer: Name of the new layer.
    :param dest_color: Color code for the new layer.
    """
    if dest_layer not in doc.layers:
        doc.layers.add(name=dest_layer, color=dest_color)
        logger.info(f"Created layer: {dest_layer}")


def change_lineweight(doc: Any, dest_layer: str, dest_lineweight: float) -> None:
    """
    Changes the lineweight for the specified layer.

    :param doc: DXF document.
    :param dest_layer: Target layer name.
    :param dest_lineweight: New lineweight value (multiplied by 10 for DXF units).
    """
    for layer in doc.layers:
        if layer.dxf.name == dest_layer:
            layer.dxf.lineweight = int(dest_lineweight * 10)
            logger.info(f"Lineweight of layer {dest_layer} changed to {dest_lineweight}")


def remove_unused_layers(doc: Any) -> None:
    """
    Removes layers that are not used in the drawing, preserving 'Defpoints' and '0'.

    :param doc: DXF document.
    """
    all_layers = [layer.dxf.name for layer in doc.layers]
    used_layers = [entity.dxf.layer for entity in doc.modelspace()]

    # Preserve essential layers
    if "Defpoints" not in used_layers:
        used_layers.append("Defpoints")
    if "0" not in used_layers:
        used_layers.append("0")

    unused = [layer for layer in all_layers if layer not in used_layers]
    for layer in unused:
        doc.layers.remove(layer)
        logger.info(f"Removed layer: {layer}")


def change_layer(msp: Any, layer: str, dest_layer: str) -> None:
    """
    Moves all entities from the original layer to the destination layer.

    :param msp: Modelspace containing the entities.
    :param layer: Original layer name.
    :param dest_layer: Destination layer name.
    """
    for entity in msp:
        if entity.dxf.layer == layer:
            entity.dxf.layer = dest_layer
            logger.info(f"Moved {entity.dxftype()} (Handle: {entity.dxf.handle}) to layer {dest_layer}")


# ===============================
#          UTILITIES
# ===============================

def explode_drawing(msp: Any) -> None:
    """
    Explodes all INSERT entities in the modelspace.

    :param msp: The drawing's modelspace.
    """
    while any(entity.dxftype() == "INSERT" for entity in msp):
        for entity in list(msp):
            if entity.dxftype() == "INSERT":
                _ = list(entity.explode())
                msp.delete_entity(entity)


# ===============================
#          BLOCKS
# ===============================

def get_removable_blocks(doc: Any) -> List[str]:
    """
    Returns a list of blocks that can be removed.
    Blocks starting with "*MODEL_SPACE" or "*PAPER_SPACE" are ignored.

    :param doc: DXF document.
    :return: List of removable block names.
    """
    removable = []
    for block_name in doc.blocks.block_names():
        if block_name.upper().startswith(("*MODEL_SPACE", "*PAPER_SPACE")):
            continue
        removable.append(block_name)
    return removable


def get_deletion_order(doc: Any, removable: List[str]) -> List[str]:
    """
    Builds a dependency graph among removable blocks and performs a topological sort.
    If block A references block B, then A must be deleted before B.

    :param doc: DXF document.
    :param removable: List of removable blocks.
    :return: Order in which the blocks should be deleted.
    :raises ValueError: If a cyclic dependency is detected.
    """
    graph = {block: set() for block in removable}
    for block in removable:
        block_def = doc.blocks.get(block)
        for entity in block_def:
            if entity.dxftype() == "INSERT":
                referenced = entity.dxf.name
                if referenced in graph:
                    graph[block].add(referenced)

    order = []
    visited = {}  # 0: not visited, 1: visiting, 2: visited

    def dfs(node: str) -> None:
        if node in visited:
            if visited[node] == 1:
                raise ValueError("Cycle detected in block dependencies.")
            return
        visited[node] = 1
        for neighbor in graph[node]:
            dfs(neighbor)
        visited[node] = 2
        order.append(node)

    for node in graph:
        if node not in visited:
            dfs(node)

    order.reverse()  # Reverse order so that outer blocks are removed first.
    return order


def purge_blocks(doc: Any) -> None:
    """
    Purges (removes) blocks from the DXF document, respecting the dependency order.

    :param doc: DXF document.
    """
    removable = get_removable_blocks(doc)
    try:
        deletion_order = get_deletion_order(doc, removable)
    except ValueError as e:
        logger.error(f"Error in dependency resolution: {e}")
        # Fallback: sort by name length as a heuristic
        deletion_order = sorted(removable, key=len, reverse=True)
        logger.info(f"Using fallback order: {deletion_order}")

    for block_name in deletion_order:
        try:
            doc.blocks.delete_block(block_name)
            logger.info(f"Deleted block: {block_name}")
        except Exception as e:
            logger.error(f"Could not delete block {block_name}: {e}")


# ===============================
#          LOGOS & EXPORT
# ===============================

def change_logos(filename: str, doc_target: Any, msp_target: Any) -> None:
    """
    Updates logos by copying entities from a source DXF file to the target document.
    Also copies linetypes and text styles that are not present in the target.

    :param filename: Path to the source DXF file containing logos.
    :param doc_target: Target DXF document.
    :param msp_target: Target document's modelspace.
    """
    doc_source = ezdxf.readfile(filename)
    msp_source = doc_source.modelspace()
    
    # Clean up the source drawing
    explode_drawing(msp_source)
    purge_blocks(doc_source)

    # Delete existing IMAGE entities in the target modelspace
    for image in msp_target.query("IMAGE"):
        msp_target.delete_entity(image)

    # Copy linetypes
    for linetype in doc_source.linetypes:
        if linetype.dxf.name not in doc_target.linetypes:
            doc_target.linetypes.add(linetype.dxf.name)

    # Copy text styles
    for text_style in doc_source.styles:
        style_name = text_style.dxf.name
        if style_name not in doc_target.styles:
            font = text_style.dxf.font if hasattr(text_style.dxf, "font") else "arial.ttf"
            doc_target.styles.add(name=style_name, font=font)

    # Copy entities
    for entity in msp_source:
        logger.info(f"Copying: {entity.dxftype()}")
        try:
            new_entity = entity.copy()
            msp_target.add_entity(new_entity)
        except Exception as e:
            logger.error(f"Skipped {entity.dxftype()} due to error: {e}")


def export_single_file(path: str, new_file: str) -> None:
    """
    Exports the DXF files from the specified directory to a new DXF file in a single dxf file.
    Each file is imported into a separate paperspace.

    :param path: Directory containing the DXF files.
    :param new_file: Path for the output DXF file.
    """
    doc = ezdxf.new()
    path = os.path.join(path, "adjusted")

    files = sorted(
    [f for f in os.listdir(path) if f.lower().endswith(".dxf")],
    key=lambda f: int(re.search(r"_(\d+)\.dxf$", f).group(1)) if re.search(r"_(\d+)\.dxf$", f) else float("inf")
)

    paperspace_dict = {}
    # Create paperspaces
    for i in range(len(files)):
        paperspace_name = f"FL{i+1}"
        paperspace_dict[paperspace_name] = doc.layouts.new(paperspace_name)

    for i, file in enumerate(files):
        file_path = os.path.join(path, file)
        doc_source = ezdxf.readfile(file_path)
        msp_source = doc_source.modelspace()
        paperspace_name = f"FL{i+1}"
        psp_target = paperspace_dict[paperspace_name]

        #Copy Layers
        for layer in doc_source.layers:
            layer_name = layer.dxf.name
            layer_color = layer.dxf.color
            layer_lineweight = layer.dxf.lineweight / 10
            create_layer(doc, layer_name, layer_color)
            change_lineweight(doc, layer_name, layer_lineweight)

        # Copy linetypes
        for linetype in doc_source.linetypes:
            if linetype.dxf.name not in doc.linetypes:
                try:
                    pattern = linetype.dxf.pattern if hasattr(linetype.dxf, "pattern") else [0.5, -0.25, 0.5]
                    description = linetype.dxf.description if hasattr(linetype.dxf, "description") else "Copied linetype"
                    doc.linetypes.add(name=linetype.dxf.name, pattern=pattern, description=description)
                except Exception as e:
                    logger.error(f"Skipped linetype {linetype.dxf.name} due to error: {e}")

        # Copy text styles
        for text_style in doc_source.styles:
            style_name = text_style.dxf.name
            if style_name not in doc.styles:
                font = text_style.dxf.font if hasattr(text_style.dxf, "font") else "arial.ttf"
                doc.styles.add(name=style_name, font=font)

        # Copy entities
        for entity in msp_source:
            logger.info(f"Copying: {entity.dxftype()}")
            try:
                new_entity = entity.copy()
                psp_target.add_entity(new_entity)
            except Exception as e:
                logger.error(f"Skipped {entity.dxftype()} due to error: {e}")

    doc.saveas(new_file)
    logger.info(f"Exported single DXF saved as: {new_file}")


def adjust_layer(logo_file: str, new_layers: str, revcloud_layers: str, path: str, path_to_save: str) -> None:
    """
    Adjusts the layers of DXF files based on a mapping defined in an Excel file.
    This process includes exploding entities, purging blocks, updating layers, and logos.

    :param logo_file: Path to the DXF file containing logos.
    :param new_layers: Path to the Excel file with layer mappings.
    :param revcloud_layers: Layers where polylines are gonna be swapped by revclouds.
    :param path: Directory containing the DXF files to be processed.
    :param path_to_save: Directory to save the adjusted DXF file.
    """
    arc_radius = 6.0                                                                #Raio do arco da nuvem de revisão
    files = [f for f in os.listdir(path) if f.lower().endswith(".dxf")]
    df = pd.read_excel(new_layers)
    os.makedirs(path_to_save, exist_ok=True)
    for i, file in enumerate(files):
        file_path = os.path.join(path, file)
        logger.info(f"Processing: {file}")

        doc = ezdxf.readfile(file_path)
        msp = doc.modelspace()

        explode_drawing(msp)
        purge_blocks(doc)
        remove_unused_layers(doc)

        layers = list_current_layers(doc)
        layers_to_modify = df[df["currentLayer"].isin(layers)].copy()

        layers_to_modify["newLayer"].fillna("fallback", inplace=True)
        layers_to_modify["colorID"].fillna(256, inplace=True)
        layers_to_modify["lineweight"].fillna(0.0, inplace=True)
        layers_to_modify["lineType"].fillna("continuous", inplace=True)

        for layer in layers_to_modify["currentLayer"]:
            row = layers_to_modify.loc[layers_to_modify["currentLayer"] == layer]
            dest_layer = row["newLayer"].values[0]
            dest_color = row["colorID"].values[0]
            dest_lineweight = row["lineweight"].values[0]
            # dest_linetype = row["lineType"].values[0]  # Currently not used

            create_layer(doc, dest_layer, dest_color)
            change_layer(msp, layer, dest_layer)
            change_lineweight(doc, dest_layer, dest_lineweight)

        for entity in msp:
            entity.dxf.color = 256       # BYLAYER
            entity.dxf.lineweight = -1     # BYLAYER
            entity.dxf.linetype = "BYLAYER"  # BYLAYER

        change_logos(logo_file, doc, msp)
        remove_unused_layers(doc)
        create_revcloud(msp, revcloud_layers, arc_radius)

        file_name = os.path.join(path_to_save, file)
        doc.saveas(file_name)
        logger.info(f"{file} processed successfully!")

def create_revcloud(msp, revcloud_layers, arc_radius):
    all_polylines = []
    all_layers = []

    for entity in list(msp):
        layer = entity.dxf.layer
        if layer in revcloud_layers:
            all_layers.append(layer)
            if entity.dxftype() == "LWPOLYLINE":
                if entity.is_closed:
                    polylines_vertices = []
                    for vertice in entity.vertices():
                        polylines_vertices.append(vertice)
                    all_polylines.append(polylines_vertices)
                    msp.delete_entity(entity)
    
    for polyline, layer in zip(all_polylines, all_layers):
        revcloud = ezdxf.revcloud.add_entity(msp, polyline, arc_radius)
        revcloud.dxf.layer = layer