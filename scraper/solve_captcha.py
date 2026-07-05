import sys
import os

try:
    import ddddocr
except ImportError:
    print("ddddocr_not_installed")
    sys.exit(1)

def main():
    if len(sys.argv) < 2:
        print("missing_image_path")
        sys.exit(1)
        
    img_path = sys.argv[1]
    if not os.path.exists(img_path):
        print("file_not_found")
        sys.exit(1)
        
    try:
        ocr = ddddocr.DdddOcr(show_ad=False)
        with open(img_path, 'rb') as f:
            img_bytes = f.read()
        res = ocr.classification(img_bytes)
        print(res.upper())
    except Exception as e:
        print(f"error_{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
