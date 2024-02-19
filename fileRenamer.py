import os
from os import listdir
import PIL
from PIL import Image
from os import mkdir

# Criar lista de elementos presentes na pasta
filenames = listdir("./Fotos")

# Ordenar os elementos alfabeticamente
filenames.sort()

# Criar lista só de elementos .jpg
jpgFiles = []
for file in filenames:
    filetype = file.split('.')[-1]
    if filetype == 'jpg':
        jpgFiles.append(file)

# Remover para poupar memória do PC
del filenames

# Redimensiona e coloca as fotos no sítio correto
for file in jpgFiles:
    origem = './Fotos/'+file
    sepFile = file.split('_')
    date = sepFile[0][2:8]
    image = Image.open(origem)
    resizedImage = image.resize((1280,720))
    if file != jpgFiles[-1]:
        
        imageName = './Fotos/'+ date + '/'+file  
    else:
        if not os.path.exists('./FSERV'):
            mkdir('./FSERV')
        fullCWD = os.getcwd()#Current Working Directory Full path
        cWD = os.path.basename(fullCWD)
        splitCWD = cWD.split('-')
        cliente = splitCWD[0].strip()
        imageName = './FSERV/Folha de Serviço.'+ cliente + '.'+ date + '.jpg'
    resizedImage.save(imageName)
