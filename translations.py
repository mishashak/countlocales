# 번역 딕셔너리
TRANSLATIONS = {
    'ko': {
        'UI_001': '언어를 선택하세요 / Select language:',
        'UI_002': '1. 한국어',
        'UI_003': '2. English',
        'UI_004': '선택 / Choice (1 or 2): ',
        'UI_005': '잘못된 선택입니다. 기본값(한국어)을 사용합니다.',
        'UI_006': '파일 경로',
        'UI_007': '보고서 파일 경로',
        'UI_008': '처리할 파일들',
        'UI_009': '총 {}개의 파일이 감지되었습니다.',
        'UI_010': '중단하려면 Ctrl+C를 누르세요.',
        'UI_011': '파일 처리 중',
        'UI_012': '시트 처리 중',
        'UI_013': '{}개 파일 완료',
        'UI_014': '보고서 생성 중...',
        'UI_015': '보고서 저장됨',
        'UI_016': '오류',
        'UI_017': '주의',
        'UI_018': '오류 발생',
        'UI_019': '계속하려면 아무 키나 누르세요...',
        'MAIN_001': '분석 방식을 선택하세요 / Select analysis type:',
        'MAIN_002': '1. 글자 수 분석 (Character Count)',
        'MAIN_003': '2. 단어 수 분석 (Word Count)',
        'MAIN_004': '선택 / Choice (1 or 2): ',
        'MAIN_005': '잘못된 선택입니다. 기본값(글자 수 분석)을 사용합니다.'
    },
    'en': {
        'UI_001': '언어를 선택하세요 / Select language:',
        'UI_002': '1. 한국어',
        'UI_003': '2. English',
        'UI_004': '선택 / Choice (1 or 2): ',
        'UI_005': 'Invalid selection. Using default (Korean).',
        'UI_006': 'file path',
        'UI_007': 'report file path',
        'UI_008': 'files to process',
        'UI_009': 'total {} files are detected.',
        'UI_010': 'Press Ctrl+C to stop.',
        'UI_011': 'processing file',
        'UI_012': 'processing sheet',
        'UI_013': '{} files are completed',
        'UI_014': 'Generating report...',
        'UI_015': 'Report saved',
        'UI_016': 'Error',
        'UI_017': 'Caution',
        'UI_018': 'caused an error',
        'UI_019': 'Press any key to continue...',
        'MAIN_001': '분석 방식을 선택하세요 / Select analysis type:',
        'MAIN_002': '1. 글자 수 분석 (Character Count)',
        'MAIN_003': '2. 단어 수 분석 (Word Count)',
        'MAIN_004': '선택 / Choice (1 or 2): ',
        'MAIN_005': 'Invalid selection. Using default (Character Count).'
    }
}

def t(key, lang):
    """번역 함수"""
    return TRANSLATIONS[lang].get(key, key)
