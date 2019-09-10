import re
import sys
from os import listdir
from excelwriter import generate_excel


def extract_float(line):  # Extract float from given line
    float_list = re.findall("\d+\.\d+", line)
    return float_list[0]


def extract_node(line):  # Extract node name from given line
    node_list = re.findall("h\d+", line)
    return node_list[0]


def get_txts(): # Finds all txt files in given directory
    txt_files = []
    files = listdir("data")
    for file in files:
        r1 = re.compile(".txt$")
        if r1.search(file):
            txt_files.append(file)
    return txt_files


txt_files = get_txts()


def generate_database():
    database = []
    total_sum = 0
    total_min = sys.float_info.max
    total_max = -1
    for file in txt_files:
        transaction_data = []
        f = open("data/" + file, "r")
        lines = f.readlines()
        for line in lines:
            if "sending" not in line:
                current_node = [extract_node(line), extract_float(line)]
                transaction_data.append(current_node)

        transaction_data = sorted(transaction_data, key=lambda x: x[1])  # Sorts the transaction data
        min_time = float(transaction_data[0][1])  #  Pick minimum tx time, local minimum
        subtract_time = min_time #  To find delays
        sum_time = 0  #  Needed for finding avg time
        max_time = -1  #  Local maximum
        is_first_node = 1
        for node in transaction_data:  #  Calculate delays
            current_time = float(node[1])
            current_delay = current_time - subtract_time
            node.append(current_delay)
            sum_time += current_delay
            if (current_delay < min_time) and not is_first_node:
                min_time = current_delay
            if current_delay > max_time:
                max_time = current_delay
            is_first_node = 0
        avg_time = sum_time / len(transaction_data)
        transaction_name = file.replace('.txt','')  #  Create tx name from file name
        transaction = [transaction_name, [avg_time, min_time, max_time], transaction_data]
        database.append(transaction)
        #  Here total statistics are calculated
        total_sum += avg_time
        if avg_time < total_min:
            total_min = avg_time
        if avg_time > total_max:
            total_max = avg_time

    total_avg = total_sum / len(txt_files)
    total_statistics = [total_avg, total_min, total_max]
    database.insert(0, total_statistics)

    return database

generate_excel(generate_database())
database = generate_database()
