from PIL import Image

img = Image.open("1024.png").convert("RGBA")

img.save(
    "app.ico",
    sizes=[
        (16, 16),
        (24, 24),
        (32, 32),
        (48, 48),
        (64, 64),
        (128, 128),
        (256, 256),
    ]
)