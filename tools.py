from datetime import datetime


def get_time_stamp():
    # Get the current date and time
    now = datetime.now()

    # Format the date and time
    formatted_date = now.strftime("%Y-%m-%d/%H:%M:%S")

    return formatted_date


if __name__ == "__main__":
    print(get_time_stamp())
