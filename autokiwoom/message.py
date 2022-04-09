############################################
# message.py
#   메세지 형태를 정의하는 파일
############################################

def outputmsg(msg_code):
    msgcode_dic = {
                    0   : ("OP_ERR_NONE",       "정상처리"),
                    -100: ("OP_ERR_LOGIN",      "사용자 정보교환 실패"),
                    -101: ("OP_ERR_CONNECT",    "서버접속 실패"),
                    -102: ("OP_ERR_VERSION",    "버전처리 실패"),
                    -106: ("OP_ERR_SOCKET_CLOSED", "통신연결 종료")
                  }
    printECode = msgcode_dic[msg_code]
    return printECode