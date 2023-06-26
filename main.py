import traceback
from SortX import CustomException, MasterIndex


try:
    mi = MasterIndex()

    #if new list arrives, update master index
    xl_folder_path = r"Need Lists\NL 1\ELECTRICAL"
    mi.update_new_list(xl_folder_path)

    folder_path = r"C:\Users\IDM252577\Desktop\Python Projects\Utility\Need List\Received Files\b1"
    mi.update_folder_link(folder_path)
    
    #if new files arrive, update master index
    df = mi.dfmaster

    mi.write_to_excel(mi.dfmaster,overwrite=True)

    mi.logger.info("Done")


except CustomException as e:
    print(e)
    print(traceback.format_exc())
