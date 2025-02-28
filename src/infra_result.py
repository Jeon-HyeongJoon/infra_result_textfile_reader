import os

class InfraResult:
    def __init__(self):
        self.code = None
        self.vulnerability_name = None
        self.system_state = None
        self.result = None

    class SystemState:
        def __init__(self):
            self.data_name = "시스템 현황"
            self.system_state = []
            
        def to_dict(self):
            return {
                'data_name': self.data_name,
                'system_state': self.system_state
            }

    class Result:
        def __init__(self):
            self.data_name = "점검결과"
            self.status = None  # 양호, 취약, 수동진단 등의 상태
            self.result = []    # 상세 설명
            
        def to_dict(self):
            return {
                'data_name': self.data_name,
                'status': self.status,
                'result': self.result
            }

    def to_dict(self):
        """객체의 데이터를 딕셔너리로 변환"""
        return {
            'code': self.code,
            'vulnerability_name': self.vulnerability_name,
            'system_state': self.system_state.to_dict() if self.system_state else None,
            'result': self.result.to_dict() if self.result else None
        }

    def __str__(self):
        """객체를 문자열로 변환 (딕셔너리 형태)"""
        return str(self.to_dict())

    @staticmethod
    def is_separator_line(line):
        """구분선 여부 확인 - '#' 로 시작하는 구분선만 제외"""
        return line.startswith('###') and line.endswith('###')

    @classmethod
    def parse_file(cls, filename, target_encoding):
        """파일을 읽어서 InfraResult 객체들의 리스트를 반환"""
        try:
            print("[*] file open: ", filename, ", encoding:", target_encoding)
            with open(filename, 'r', encoding=target_encoding) as file:
                lines = [line.strip() for line in file.readlines()]
                return cls.parse_vulnerabilities(lines)
        except FileNotFoundError:
            print(f"파일을 찾을 수 없습니다: {filename}")
            return None
        except Exception as e:
            print(f"파일 처리 중 오류 발생: {e}")
            return None

    @classmethod
    def parse_vulnerabilities(cls, lines):
        results = []
        current_lines = []
        in_final_section = False
        
        for i, line in enumerate(lines):
            # 파일 시작 부분의 불필요한 내용 건너뛰기
            if not current_lines and not line.startswith('[W-'):
                continue
                
            # 마지막 취약점 이후의 전체 진단 결과 섹션 체크
            if line.startswith('[Unicode]') or line.startswith('[System Access]'):
                in_final_section = True
                if current_lines:  # 마지막 취약점 처리
                    results.append(cls.from_txt(current_lines))
                break
                
            if in_final_section:
                continue
                
            # 새로운 취약점 시작 또는 파일의 마지막 줄
            if (line.startswith('[W-') and current_lines) or i == len(lines) - 1:
                if i == len(lines) - 1 and not line.startswith('[W-'): 
                    current_lines.append(line)
                results.append(cls.from_txt(current_lines))
                current_lines = []
                if line.startswith('[W-'):
                    current_lines.append(line)
                continue
                
            current_lines.append(line)
        
        # 남은 마지막 취약점 처리
        if current_lines and not in_final_section:
            results.append(cls.from_txt(current_lines))
        
        return [result for result in results if result and result.code]

    @classmethod
    def from_txt(cls, content):
        """취약점 내용을 InfraResult 객체로 변환"""
        infra_result = cls()
        infra_result.system_state = cls.SystemState()
        infra_result.result = cls.Result()
        
        current_section = None
        system_state_lines = []
        
        for i, line in enumerate(content):
            line = line.strip()
            
            if not line:  # 빈 줄만 건너뛰기
                continue
                
            if cls.is_separator_line(line):  # '#' 구분선 건너뛰기
                continue
                
            if line.startswith('[W-'):
                infra_result.code = line.split(']')[0].strip('[')
                infra_result.vulnerability_name = line.split(']')[1].strip() if ']' in line else None
                continue
                
            if '■ 시스템현황' in line:
                current_section = 'system_state'
                continue
            elif '■ 점검결과' in line:
                current_section = 'result'
                if ':' in line:
                    result_status = line.split(':')[1].strip()
                    infra_result.result.status = result_status
                continue
                
            if current_section == 'system_state':
                if '■ 점검결과' not in line:
                    if line.startswith('※') and system_state_lines:
                        system_state_lines.append('')
                    system_state_lines.append(line)
                    
            elif current_section == 'result':
                if i < len(content) - 1 and not content[i + 1].startswith('[W-'):
                    if not line.startswith('[참고]'):
                        infra_result.result.result.append(line)
        
        # 리스트를 문자열로 변환
        if system_state_lines:
            infra_result.system_state.system_state = '\n'.join(system_state_lines)
        if infra_result.result.result:
            infra_result.result.result = '\n'.join(infra_result.result.result)
        
        return infra_result

def parse_result_file(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            content = file.readlines()
            return InfraResult.from_txt(content)
    except Exception as e:
        print(f"파일 처리 중 오류 발생: {e}")
        return None

