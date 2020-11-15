import datetime


def main():
    date = input("Date: ")
    delimiter = input("Delimiter: ")
    checkDate(date, delimiter)


def checkDate(date, delimiter):
    date_arr = date.split(delimiter)
    correctDate = None
    try:
        newDate = datetime.date.fromisoformat(
            f"{date_arr[2]}-{date_arr[0]}-{date_arr[1]}")
        correctDate = True
    except ValueError:
        correctDate = False
    print(str(correctDate))


if __name__ == "__main__":
    main()
