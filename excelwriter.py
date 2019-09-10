import tempfile
import sys
import xlrd
import xlwt
import os.path


def generate_style(colour_index, style_code):  # Create style for the cell, mainly for choosing background colour
    colours = ["light_green", "light_turquoise", "light_yellow"]
    title_colours = ["green", "turquoise", "yellow"]
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map[colours[colour_index]]
    style.pattern = pattern
    if style_code == 1: #  For title
        pattern.pattern_fore_colour = xlwt.Style.colour_map[title_colours[colour_index]]
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.wrap = 1
        style.alignment = alignment
    if style_code == 2: #  For local statistics
        pattern.pattern_fore_colour = xlwt.Style.colour_map[title_colours[colour_index]]
    return style

def generate_excel(database):
    colour_index = 0  # For generate_style
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("main")

    tx_count = len(database) - 1  # 1 is for total statistics
    node_count = len(database[1][2])


    for i in range(tx_count):
        local_statistics_style = generate_style(colour_index, 2)
        title_style = generate_style(colour_index, 1)
        style = generate_style(colour_index, 0)
        colour_index = (colour_index + 1) % 3

        sheet.write_merge(0, 1, 3*i, 3*i+2, "t" + str(i+1) + "\n" + str(database[i+1][0]), title_style)
        sheet.col(3*i+1).width = 8500  # Set the width of float holding columns so they can fit
        sheet.col(3*i+2).width = 8500

        for j in range(node_count):

            if j%2 == 0:
                sheet.write(j+2, 3*i, database[i+1][2][j][0])
                sheet.write(j+2, 3*i+1, database[i+1][2][j][1])
                sheet.write(j+2, 3*i+2, database[i+1][2][j][2])
            else:
                sheet.write(j+2, 3*i, database[i+1][2][j][0], style)
                sheet.write(j+2, 3*i+1, database[i+1][2][j][1], style)
                sheet.write(j+2, 3*i+2, database[i+1][2][j][2], style)
        sheet.write(node_count+2, 3*i, "AVERAGE", local_statistics_style)
        sheet.write(node_count+2, 3*i+1, "", local_statistics_style)
        sheet.write(node_count+2, 3*i+2, database[i+1][1][0], local_statistics_style)
        sheet.write(node_count+3, 3*i, "MINIMUM", local_statistics_style)
        sheet.write(node_count+3, 3*i+1, "", local_statistics_style)
        sheet.write(node_count+3, 3*i+2, database[i+1][1][1], local_statistics_style)
        sheet.write(node_count+4, 3*i, "MAXIMUM", local_statistics_style)
        sheet.write(node_count+4, 3*i+1, "", local_statistics_style)
        sheet.write(node_count+4, 3*i+2, database[i+1][1][2], local_statistics_style)

    sheet.write(node_count + 8, 1, "Total Average: ")
    sheet.write(node_count + 9, 1, "Total Minimum: ")
    sheet.write(node_count + 10, 1, "Total Maximum: ")
    sheet.write(node_count + 8, 2, str(database[0][0]))
    sheet.write(node_count + 9, 2, str(database[0][1]))
    sheet.write(node_count + 10, 2, str(database[0][2]))


    workbook.save("kshell-data.xls")

