# -*- coding: utf-8 -*-
# author lis
# http://www.lisz-works.com

# OK=True
def chkString(text):
    if type(text) != str:
        return False
    return len(str(text)) > 0

# OK=True
def chkList(l):
    if type(l) != list:
        return False
    return len(l) > 0

if __name__ == '__main__':
    print(str(chkString("")))
    print(str(chkString("a")))
    print(str(chkString("0123456789")))
