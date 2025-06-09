import pyautogui
from PIL import ImageGrab


#VER POSIÇÃO DO MOUSE

try:
    while True:
        x, y = pyautogui.position()
        positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
        
        # Captura a tela
        image = ImageGrab.grab()
        
        # Verifica se as coordenadas estão dentro dos limites da imagem
        if 0 <= x < image.width and 0 <= y < image.height:
            color = image.getpixel((x, y))
            colorStr = ' Cor: ' + str(color)
        else:
            colorStr = ' Cor: fora dos limites'
        
        print(positionStr + colorStr, end='')
        print('\b' * (len(positionStr) + len(colorStr)), end='', flush=True)
except KeyboardInterrupt:
    print('\n')