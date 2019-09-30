import string
string.punctuation


def area(length, width):
    """"
    Compute the area of a rectangle

    Parameters:
    length, width (float)
    Returns:
    area (float)
    """
    result = length * width
    return result


def main():
    bedroom = area(11, 9)  # first function call
    kitchen = area(12, 7)  # second function call
    family = area(12, 12)  # third function call
    print(bedroom,"\n", kitchen, family)

main()