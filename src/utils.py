def normalize_key(key):
    try:
        return key.char
    except AttributeError:
        return str(key).replace("Key.", "")