import cv2
import pytesseract
import argparse
from PIL import Image
from docx import Document
import os
import re

def preprocess_image(image_path, method):
    image = cv2.imread(image_path)

    if method == "gray":
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    elif method == "thresh":
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, image = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
    elif method == "blur":
        image = cv2.GaussianBlur(image, (5, 5), 0)

    temp_path = "temp_processed.png"
    cv2.imwrite(temp_path, image)
    return temp_path

def normalize(text):
    return re.sub(r'\W+', '', text.lower())

def calculate_accuracy(recognized, expected):
    norm_recognized = normalize(recognized)
    norm_expected = normalize(expected)

    correct = sum(1 for a, b in zip(norm_recognized, norm_expected) if a == b)
    total = max(len(norm_expected), 1)
    return correct / total

def main():
    parser = argparse.ArgumentParser(description="OCR ảnh tiếng Việt và lưu vào file DOCX")
    parser.add_argument("-i", "--image", required=True, help="Đường dẫn tới ảnh đầu vào")
    parser.add_argument("-p", "--preprocess", type=str, default="thresh",
                        choices=["gray", "thresh", "blur"], help="Phương pháp xử lý ảnh")
    parser.add_argument("--expected", help="Đường dẫn tới file văn bản gốc để so sánh (tùy chọn)")

    args = parser.parse_args()

    print("🖼️ Đang xử lý ảnh...")
    processed_image_path = preprocess_image(args.image, args.preprocess)

    print("🔍 Đang nhận dạng văn bản...")
    text = pytesseract.image_to_string(Image.open(processed_image_path), lang="vie")

    print("\n--- Recognized Text ---")
    print(text)

    # Ghi ra file DOCX
    print("💾 Đang lưu vào file DOCX...")
    doc = Document()
    doc.add_paragraph(text)
    doc.save("ketqua.docx")
    print("✅ Đã lưu vào 'ketqua.docx'")

    # Nếu có cung cấp file expected thì so sánh
    if args.expected:
        if os.path.exists(args.expected):
            with open(args.expected, "r", encoding="utf-8") as f:
                expected_text = f.read()
            accuracy = calculate_accuracy(text, expected_text)
            print(f"\n🎯 Độ chính xác OCR (sau chuẩn hóa): {accuracy:.2%}")
        else:
            print("⚠️ Không tìm thấy file expected để so sánh.")

    # Xóa ảnh tạm nếu cần
    if os.path.exists(processed_image_path):
        os.remove(processed_image_path)

if __name__ == "__main__":
    main()
