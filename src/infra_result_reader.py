import os
import pandas as pd
from infra_result import InfraResult

def convert_to_dataframe(results):
    """InfraResult 객체들을 DataFrame으로 변환"""
    data = []
    for i, result in enumerate(results):
        # 디버깅을 위한 출력
        print(f"Processing result {i+1}: {result.code}")
        
        # 리스트인 경우 문자열로 변환
        system_state = result.system_state.system_state if result.system_state else None
        if isinstance(system_state, list):
            system_state = '\n'.join(system_state)
            
        # result_detail = result.result.result if result.result else None
        # if isinstance(result_detail, list):
        #     result_detail = '\n'.join(result_detail)
            
        row = {
            '취약점 코드': result.code,
            '취약점 이름': result.vulnerability_name,
            '점검 결과': result.result.status if result.result else None,
            '시스템 현황': system_state,
            # '상세 설명': result_detail
        }
        data.append(row)
        
    print(f"\nTotal rows to be converted: {len(data)}")
    
    df = pd.DataFrame(data)
    return df

def save_to_excel(results, output_path):
    df = convert_to_dataframe(results)
    
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    
    df.to_excel(writer, sheet_name='취약점 분석 결과', index=False, startrow=0)
    
    # 워크북과 워크시트 객체 가져오기
    workbook = writer.book
    worksheet = writer.sheets['취약점 분석 결과']
    
    # 셀 포맷 설정
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#D9E1F2',
        'border': 1,
        'font_size': 9  # 헤더 글자 크기
    })
    
    # 기본 정렬 포맷맷
    cell_format = workbook.add_format({
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'left',
        'border': 1,
        'font_size': 9  # 데이터 글자 크기
    })

    # 가운데 정렬 포맷
    cell_format_center = workbook.add_format({
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'border': 1,
        'font_size': 9
    })
    
    # 열 너비 설정
    # worksheet.set_column('A:A', 15)  # 취약점 코드
    # worksheet.set_column('B:B', 30)  # 취약점 이름
    # worksheet.set_column('C:C', 15)  # 점검 결과
    # worksheet.set_column('D:D', 50)  # 시스템 현황
    # worksheet.set_column('E:E', 50)  # 상세 설명 - 제외됨
    
    # 헤더 포맷 적용
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    # 데이터 포맷 적용
    for row_num in range(len(df)):
        # 현재 행의 D열(점검 결과) 텍스트 길이에 따라 행 높이 계산
        d_column_text = str(df.iloc[row_num, 3])  # D열은 인덱스 3
        # 줄바꿈 문자로 분리하여 라인 수 계산
        num_lines = len(d_column_text.split('\n'))
        # 기본 행 높이 15, 줄당 15픽셀 추가
        row_height = max(15, num_lines * 15)
        
        # 행 높이 설정
        worksheet.set_row(row_num + 1, row_height)
        
        for col_num in range(len(df.columns)):
            value = df.iloc[row_num, col_num]
            if pd.isna(value):  # None 값 처리
                value = ''

            if col_num in [0, 2]:  # A열과 C열
                worksheet.write(row_num + 1, col_num, value, cell_format_center)
            else:
                worksheet.write(row_num + 1, col_num, value, cell_format)
    
    # 파일 저장
    writer.close()
    
    print(f"결과가 {output_path}에 저장되었습니다.")

def get_result_files(target_dir):
    if not os.path.exists(target_dir):
        print(f"오류: '{target_dir}' 폴더를 찾을 수 없습니다.")
        return []
    
    # result 폴더 내의 모든 파일 목록 가져오기
    files = [f for f in os.listdir(target_dir) if os.path.isfile(os.path.join(target_dir, f))]
    return files

def make_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"'{path}' 폴더가 생성되었습니다.")

def main():
    # 현재 경로를 src 폴더의 상위 경로로 설정
    current_path = os.path.dirname(os.path.dirname(os.getcwd()))
    target_dir = os.path.join(current_path, "target")
    if not os.path.exists(target_dir):
        raise RuntimeError("target 폴더에 진단결과 txt 파일을 넣어주세요.")
    
    # 결과 파일 경로 설정
    output_file = os.path.join(current_path, "result.xlsx")
    
    # result 폴더 내의 모든 파일 목록 가져오기
    result_files = get_result_files(target_dir)
    
    if not result_files:
        print("처리할 파일이 없습니다.")
        return
    
    print("\n=== 처리할 파일 목록 ===")
    for i, file in enumerate(result_files, 1):
        print(f"{i}. {file}")
    
    total_processed = 0
    total_errors = 0
    
    # ExcelWriter 객체 생성
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    workbook = writer.book
    
    print("\n=== 파일 처리 시작 ===")
    for input_file_name in result_files:
        input_file = os.path.join(target_dir, input_file_name)
        sheet_name = os.path.splitext(input_file_name)[0].split('(')[0]
        
        print(f"\n처리 중: {input_file_name}")
        try:
            # 결과 파일 파싱
            results = InfraResult.parse_file(input_file)
            
            if results:
                # DataFrame 생성
                df = convert_to_dataframe(results)
                
                # 시트에 데이터 쓰기
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                
                # 워크시트 객체 가져오기
                worksheet = writer.sheets[sheet_name]
                
                # 셀 포맷 설정
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': '#D9E1F2',
                    'border': 1,
                    'font_size': 9
                })
                
                cell_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'left',
                    'border': 1,
                    'font_size': 9
                })
                
                cell_format_center = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'center',
                    'border': 1,
                    'font_size': 9
                })
                
                # 열 너비 설정
                # worksheet.set_column('A:A', 15)  # 취약점 코드
                # worksheet.set_column('B:B', 30)  # 취약점 이름
                # worksheet.set_column('C:C', 15)  # 점검 결과
                # worksheet.set_column('D:D', 50)  # 시스템 현황
                
                # 헤더 포맷 적용
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # 데이터 포맷 적용
                for row_num in range(len(df)):
                    # 현재 행의 D열(점검 결과) 텍스트 길이에 따라 행 높이 계산
                    d_column_text = str(df.iloc[row_num, 3])
                    num_lines = len(d_column_text.split('\n'))
                    row_height = max(15, num_lines * 15)
                    
                    # 행 높이 설정
                    worksheet.set_row(row_num + 1, row_height)
                    
                    for col_num in range(len(df.columns)):
                        value = df.iloc[row_num, col_num]
                        if pd.isna(value):
                            value = ''
                        
                        if col_num in [0, 2]:  # A열과 C열
                            worksheet.write(row_num + 1, col_num, value, cell_format_center)
                        else:
                            worksheet.write(row_num + 1, col_num, value, cell_format)
                
                print(f"성공: {input_file_name} -> {len(results)}건의 결과가 변환되었습니다.")
                total_processed += 1
            else:
                print(f"실패: {input_file_name} - 파일을 읽을 수 없거나 변환 중 오류가 발생했습니다.")
                total_errors += 1
                
        except Exception as e:
            print(f"오류: {input_file_name} - {str(e)}")
            total_errors += 1
    
    # 파일 저장
    writer.close()
    
    print(f"\n=== 처리 완료 ===")
    print(f"총 처리된 파일: {total_processed}개")
    print(f"실패한 파일: {total_errors}개")
    print(f"결과가 {output_file}에 저장되었습니다.")

if __name__ == "__main__":
    main()