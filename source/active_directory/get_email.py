import win32api
import win32net
from pyad import aduser


def get_user_email() -> str:
    dc_name = win32net.NetGetAnyDCName()
    user = win32api.GetUserName()
    user_info = win32net.NetUserGetInfo(dc_name, user, 4)  # last number - level of access
    cn = user_info["full_name"]
    user = aduser.ADUser.from_cn(cn)
    email = user.get_attribute("mail", always_return_list=False)
    return email

print(get_user_email())
