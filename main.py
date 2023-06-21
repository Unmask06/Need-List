import traceback
from SortX import CustomException, MasterIndex


try:
    mi = MasterIndex()

    folder_path = r"Need Lists\NL 1\ELECTRICAL"

    mi.merge_excel(folder_path)

    df = mi.dfmaster

    mi.update_master_index()


except CustomException as e:
    print(e)
    print(traceback.format_exc())