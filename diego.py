import dxfapi

logoFile = "G:/Meu Drive/Datacamp/Python/finalversion/GERDAU.dxf"
newlayers = "G:/Meu Drive/Datacamp/Python/editdxf/newlayers.xlsx"
new_file = "G:/Meu Drive/Datacamp/Python/consolidado2.dxf"
path = "G:/Meu Drive/Datacamp/Python/editdxf/files"

dxfapi.adjust_layer(logoFile, newlayers, path)
dxfapi.export_gerdau(path, new_file)