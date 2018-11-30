def filter_characters(string):
    # string = string.replace(" : ", "")
    # string = string.replace(": ", "")
    # string = string.replace(" :", "")
    if not (string.lower().__contains__('am') or string.lower().__contains__('pm')):
        string = string.replace(":", "")

    # string = string.replace(" , ", "")
    # string = string.replace(", ", "")
    # string = string.replace(" ,", "")
    string = string.replace(",", "")
    string = string.replace("_", " ")
    string = string.replace("-", " ")

    string = string.replace(chr(10),"")
    if(string[-1]==" "):
        string= string[:-1]
    if(string[0]==" "):
        string= string[1:]
    return string
