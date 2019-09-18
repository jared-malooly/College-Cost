from xlrd import *
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np

def main():
    book = open_workbook('costData.xls')
    sheet = book.sheet_by_index(0)
    dates = extract_dates(sheet)

    # Read through Excel sheet to grab data for PUBLIC schooling
    public_tuition, \
    public_housing, \
    public_board = extract_costs(sheet, 72, 134)

    # Read through Excel sheet to grab data for PRIVATE schooling
    private_tuition, \
    private_housing, \
    private_board = extract_costs(sheet, 137, 199)

    public_data = (public_tuition,
            public_housing,
            public_board,
            'Average Yearly Cost to Attend a Public College')

    private_data = (private_tuition,
              private_housing,
              private_board,
              'Average Yearly Cost to Attend a Private College')

    # Two plots for comparison
    plot_data_side(dates, public_data, private_data)

    # Create PUBLIC plot individually for download
    plot_data_indiv(dates,
              public_tuition,
              public_housing,
              public_board,
              'Average Yearly Cost to Attend a Public College')

    # Create PRIVATE plot individually for download
    plot_data_indiv(dates,
              private_tuition,
              private_housing,
              private_board,
              'Average Yearly Cost to Attend a Private College')

def plot_data_side(dates, public_data, private_data):
    # Datasets = (tuition, housing, board, title)
    x = range(dates[0], dates[0] + len(dates))
    fig = plt.figure()
    sns.set(palette='pastel')
    plt.subplots_adjust(hspace = .5)

    public_graph = fig.add_subplot(211)
    private_graph = fig.add_subplot(212)

    public_graph.title.set_text(public_data[3])
    public_graph.set_ylim(0,40000)
    private_graph.title.set_text(private_data[3])
    private_graph.set_ylim(0,40000)

    y1 = [public_data[0], public_data[1], public_data[2]]
    y2 = [private_data[0], public_data[1], public_data[2]]

    private_graph.stackplot(x, y2)
    public_graph.stackplot(x, y1)

    plt.show()

def plot_data_indiv(dates, tuition, housing, board, title):
    '''
    Using information from excel data, plot a stacked line graph displaying total cost
    and what makes up the cost by area.
    :param dates: 1963 - 2018
    :param tuition: Cost of tuition, by year
    :param housing: Cost of housing, by year
    :param board: Cost of board, by year
    :param title: This function is reusable for different datasets
                    ex: Public colleges vs private
    '''
    x = range(dates[0], dates[0] + len(dates))
    y = [tuition, housing, board]

    sns.set(palette='pastel')
    plt.stackplot(x, y, labels=['Tuition', 'Housing', 'Board'])
    plt.legend(loc='upper left')
    plt.title(title)
    plt.xticks(np.arange(dates[0], dates[-1] + 1, 17))
    plt.show()


def extract_dates(sheet):
    '''
    Append the data's availible dates to a list for easy comparison
    :param sheet: excel sheet object (data)
    :return list of dates:
    '''

    dates = []
    for row in range(72, 134):
        cell = sheet.cell(row,0)
        dates.append(str(cell).strip(".'text: mpty"))
    while '' in dates:
        dates.remove('')

    for index in range(len(dates)):
        dates[index] = int(dates[index].split('-')[0])

    return dates

def extract_costs(sheet, start_row, end_row):
    '''

    :param sheet: excel sheet object (data)
    :param start_row: starting position of desired data
    :param end_row: ending position of desired data
    :return:
    '''
    tuition = []
    housing = []
    board = []
    for column in range(4, 13, 3):
        for row in range(start_row, end_row):
            cell = sheet.cell(row, column)
            if column == 4:
                tuition.append(str(cell).strip(".'tex: mptynumber"))
            elif column == 7:
                housing.append(str(cell).strip(".'tex: mptynumber"))
            else:
                board.append(str(cell).strip(".'tex: mptynumber"))

        # Parse out blank cells
        while '' in tuition:
            tuition.remove('')
        while '' in housing:
            housing.remove('')
        while '' in board:
            board.remove('')

    # Convert all values to float
    for index in range(len(tuition)):
        tuition[index] = float(tuition[index])
    for index in range(len(housing)):
        housing[index] = float(housing[index])
    for index in range(len(board)):
        board[index] = float(board[index])

    return tuition, housing, board


main()