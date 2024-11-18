# %%
def 행추가(n):
    for i in range(1, n):
        hwp.HAction.Run("TableAppendCol")
        next

# %%
# 표만들기
from win32com.client import Dispatch

def add_table(hwp, rows, cols, row_height=0, col_width=0):
    """
    HWP 문서에 테이블을 생성하는 함수

    Parameters:
        hwp (object): HWP COM 객체
        rows (int): 테이블의 행 개수
        cols (int): 테이블의 열 개수
        row_height (float, optional): 행 높이 (기본값: 0)
        col_width (float, optional): 열 너비 (기본값: 0)

    Returns:
        bool: 테이블 생성 성공 여부
    """
    try:
        # 기본 동작 설정
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
        htable = hwp.HParameterSet.HTableCreation

        # 행과 열 설정
        htable.Rows = rows
        htable.Cols = cols

        # 열 너비 설정
        if col_width == 0:
            htable.WidthType = 1  # 열 너비 자동
        else:
            htable.WidthType = 2  # 열 너비 사용자 정의
            htable.CreateItemArray("ColWidth", cols)
            for i in range(cols):
                htable.ColWidth.SetItem(i, hwp.MiliToHwpUnit(col_width))

        # 행 높이 설정
        if row_height == 0:
            htable.HeightType = 0  # 행 높이 자동
        else:
            htable.HeightType = 1  # 행 높이 사용자 정의
            htable.CreateItemArray("RowHeight", rows)
            for i in range(rows):
                htable.RowHeight.SetItem(i, hwp.MiliToHwpUnit(row_height))

        # 기타 테이블 속성 설정
        htable.TableProperties.TreatAsChar = 1
        htable.TableProperties.HorzOffset = hwp.MiliToHwpUnit(0)
        htable.TableProperties.VertOffset = hwp.MiliToHwpUnit(0)
        htable.TableTemplate.CreateMode = 0
        htable.TableProperties.Width = 53294

        # 테이블 생성 실행
        hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

        return True  # 성공
    except Exception as e:
        print(f"테이블 생성 중 오류 발생: {e}")
        return False  # 실패


# %%
def append_hwp(filename):
    """
    문서 끼워넣기
    """
    hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
    hwp.HParameterSet.HInsertFile.KeepSection = 0
    hwp.HParameterSet.HInsertFile.KeepCharshape = 0
    hwp.HParameterSet.HInsertFile.KeepParashape = 0
    hwp.HParameterSet.HInsertFile.KeepStyle = 0
    hwp.HParameterSet.HInsertFile.filename = filename
    hwp.HAction.Execute("InsertFile", hwp.HParameterSet.HInsertFile.HSet)


