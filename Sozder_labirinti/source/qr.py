# QR-код: ойын сілтемесі (стенд пен қосымша үшін)
import qrcode, sys
url = sys.argv[1] if len(sys.argv) > 1 else "https://claude.ai/artifact/NWRfvmZBJcKwFvs5nTSzuU"
img = qrcode.make(url, box_size=20, border=2, error_correction=qrcode.constants.ERROR_CORRECT_M)
img.save("figs/qr.png")
print("QR:", url)
