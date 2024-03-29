from datetime import datetime

# Get the current date and time
now = datetime.now()

# Print the current date and time
print(f"The current date and time is {now}.")


def get_time_stamp():
    date = datetime.strptime(datetime.now, "%Y-%m-%dT%H;%M:%S.%fZ")
    formatted_date = date.strftime("%Y-%m-%d/%H:%M:%S")
    return formatted_date
