import os
from os import listdir
import PIL
from PIL import Image
from os import mkdir

def getJPEG():
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
  return jpgFiles

def resizeFiles(jpgFiles):
  resizedImages = []
  for file in jpgFiles:
    origin = './Fotos/'+str(file)
    sepFile = file.split('_')
    date = sepFile[0][2:8]
    image = Image.open(origin)
    resizedImage = image.resize((1280,720))
    resizedImages.append(resizedImage)
  return resizedImages, date

def imageAllocator(resizedImages, date):
  counter = 0
  for image in resizedImages:
    counter += 1
    if image !=resizedImages[-1]:
      imageDir = './Fotos/'+ str(date) + '/'+ str(counter)
    else:
      if not os.path.exists('./FSERV'):
        mkdir('./FSERV')
      fullCWD = os.getcwd()#Current Working Directory Full path
      cWD = os.path.basename(fullCWD)
      splitCWD = cWD.split('-')
      cliente = splitCWD[0].strip()
      imageDir = './FSERV/Folha de Serviço.'+ cliente + '.'+ date + '.jpg'
    image.save(imageDir)

def main():
    jpgFiles = getJPEG()
    resizedImages, date = resizeFiles(jpgFiles)
    imageAllocator(resizedImages, date)

main()
