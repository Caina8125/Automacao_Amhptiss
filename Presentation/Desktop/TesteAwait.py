import asyncio
from Presentation.Desktop.Gif import ImageLabel


async def chamada_assincrona(self,):
    resultado = await ImageLabel.iniciarGif()
    

# Executando o loop de evento assíncrono para chamar a função assíncrona
# asyncio.run(chamada_assincrona())
