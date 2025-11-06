import os
import sys
from count_chars import main as count_chars_main
from count_words import main as count_words_main
from translations import t

def select_language():
    """언어 선택 함수"""
    print(t('UI_001', 'ko'))
    print(t('UI_002', 'ko'))
    print(t('UI_003', 'ko'))
    
    try:
        choice = input(t('UI_004', 'ko')).strip()
        if choice == '1':
            return 'ko'
        elif choice == '2':
            return 'en'
        else:
            print(t('UI_005', 'ko'))
            return 'ko'  # 기본값
    except:
        print(t('UI_005', 'ko'))
        return 'ko'  # 기본값

def select_analysis_type(lang):
    """분석 방식 선택 함수"""
    print(t('MAIN_001', lang))
    print(t('MAIN_002', lang))
    print(t('MAIN_003', lang))
    
    try:
        choice = input(t('MAIN_004', lang)).strip()
        if choice == '1':
            return 'chars'
        elif choice == '2':
            return 'words'
        else:
            print(t('MAIN_005', lang))
            return 'chars'  # 기본값
    except:
        print(t('MAIN_005', lang))
        return 'chars'  # 기본값


def main():
    # 언어 선택
    current_language = select_language()
    
    # 분석 방식 선택
    analysis_type = select_analysis_type(current_language)
    
    # 선택된 분석 방식에 따라 실행
    if analysis_type == 'chars':
        count_chars_main(current_language)
    elif analysis_type == 'words':
        count_words_main(current_language)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
        input("Press any key to continue...")
