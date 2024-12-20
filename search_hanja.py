"""
main 함수. 폴더를 순회하며 hwp, pdf, xlsx 파일에 한해 개인정보를 수집합니다
"""

import os


from module.processing_folder import processing_folder


def main():
    """개인정보를 추출합니다"""
    print("-"*24)
    print("\n>>>>>>문서 내 한자 검색기<<<<<<\n")
    print("-"*24)
    input_path = input("작업할 폴더 경로를 입력하세요(종료는 0을 입력) : ").strip()

    if input_path == '0':
        return 0

    output_path = input(
        "엑셀파일 경로를 입력하세요(확장자포함. 파일이 존재하지 않을 경우 새로 생성) : ").strip()

    if not os.path.isdir(input_path):
        print("입력 폴더의 경로를 다시 확인하세요")
        return main()
    print("\n")
    processing_folder(input_path, output_path)
    print("-"*24)
    print(f"{output_path}에 검색 결과 목록이 생성되었습니다.")
    print("\n~~~모든 작업이 완료되었습니다~~~")

    return main()


if __name__ == "__main__":
    main()
