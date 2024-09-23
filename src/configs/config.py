import socket

ORG_PDF_PATH = "D:/projects/pdf2excel/others/org_sample_file/(ESSENTIEL) 20230307 - WOMEN FALL W2023 - PURCHASE ORDER DROP 2 - HANGZHOU GLORIA IMP.& EXP. CO., LTD - PO-0006742 - REVISED 20230515.pdf"
BOOL_MERGE_HEAD_ROW = True
BOOL_MERGE_HEAD_COL = True

TITLE = "PDF2Excel"
VERSION = "0.0.1"


# ip地址相关操作
def get_local_ip():
    try:
        # 创建一个socket对象
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        # 不需要真正发送数据，所以使用一个不存在的地址
        s.connect(("10.255.255.255", 1))
        # 获取本地IP地址
        IP = s.getsockname()[0]
    except Exception as e:
        # 在出现异常时，可能无法获取IP，比如网络未连接
        IP = "127.0.0.1"
    finally:
        # 关闭socket连接
        s.close()
    return IP


IP = get_local_ip()
