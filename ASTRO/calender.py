import datetime


def get_yeardatetime_list():

    base_date = datetime.date(2019, 1, 1)

    current_date = datetime.datetime.now()

    day_size = current_date.date() - base_date

    increasing_elem = day_size.days

    yeardatetime_list = []

    for i in range(1, increasing_elem+1):
        yeardatetime_list.append(base_date + datetime.timedelta(days=i))

    return yeardatetime_list


if __name__ == '__main__':
    list_data = get_yeardatetime_list()
    for date in list_data:
        print(type(date.year))