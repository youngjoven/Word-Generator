"""
YAML 파일에서 Word 문서를 생성하는 메인 스크립트

사용법:
    python create_document.py

생성 파일:
    output.docx (또는 사용자 지정 파일명)
"""
import yaml
from document_generator import DocumentGenerator


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("Word 산출물 문서 생성 프로그램")
    print("=" * 60)
    print()

    # YAML 파일 읽기
    input_file = 'sample_data.yaml'
    output_file = 'output.docx'

    print(f"📄 {input_file} 파일을 읽는 중...")
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        print("✅ 파일 읽기 완료")
    except FileNotFoundError:
        print(f"❌ 오류: {input_file} 파일을 찾을 수 없습니다.")
        return
    except yaml.YAMLError as e:
        print(f"❌ YAML 파싱 오류: {e}")
        return
    print()

    # Word 문서 생성
    print("📝 Word 문서를 생성하는 중...")
    try:
        generator = DocumentGenerator()
        generator.generate_document(data, output_file)
        print("✅ Word 문서 생성 완료")
    except Exception as e:
        print(f"❌ 문서 생성 오류: {e}")
        return
    print()

    print("=" * 60)
    print(f"🎉 {output_file} 파일이 생성되었습니다!")
    print("=" * 60)
    print()
    print("💡 참고사항:")
    print("   - 문서를 열면 페이지 번호가 자동으로 업데이트됩니다")
    print("   - 목차를 클릭하면 해당 섹션으로 이동합니다")
    print()


if __name__ == "__main__":
    main()
