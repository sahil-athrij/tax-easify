def about():
    f = open('about.txt','r')
    a = f.read()
    f.close()
    return a

def helper():
    f = open('help.txt','r')
    a = f.read()
    f.close()
    return a
