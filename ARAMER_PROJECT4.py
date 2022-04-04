import openpyxl
import numbers
from map import us_state_to_abbrev


# this function is the main function that uses all the other functions I defined in the code
def main():
    pop_data_sheet = open_worksheet("countyPopChange2020-2021.xlsx")
    should_get_losses_variable = should_get_losses()
    process_data(pop_data_sheet, should_get_losses_variable)


# this function basically opens up the Excel document and returns its data
def open_worksheet(filename):
    income_excel = openpyxl.load_workbook(filename)
    data_sheet = income_excel.active
    return data_sheet


# this function asks the user which result they would like to see (loss or gain in population) and returns a true or
# false value depending on the user's response
def should_get_losses():
    show_losses = input("Would you like to get the counties that lost population?")
    no = ["no", "nah", "nope"]  # this is a list of possible answers the user could give back that are equivalent to no
    yes = ["yes", "ya", "yeah"]  # this is a list of possible answers the user could give back that are equivalent to yes
    if show_losses in no:
        return False
    if show_losses in yes:
        return True


# this is the big function that does most of the mathematical work, and prints the results.
def process_data(data_sheet, show_losses):
    for row in data_sheet.rows:
        state_name = row[5]
        state_name_value = state_name.value
        county_name = row[6]
        county_name_value = county_name.value
        pop_estimate2021 = row[9]
        n_pop_chg2021 = row[11]
        pop_estimate2021_value = pop_estimate2021.value
        n_pop_chg2021_value = n_pop_chg2021.value
        if not isinstance(pop_estimate2021_value, numbers.Number):
            continue
        if not isinstance(n_pop_chg2021_value, numbers.Number):
            continue
        if state_name.value not in us_state_to_abbrev:
            continue
        pop_change = pop_estimate2021_value - n_pop_chg2021_value
        pop_change_percent = n_pop_chg2021_value / pop_change
        pop_change_percent = pop_change_percent * 100
        if show_losses:
            if pop_change_percent < -2.0:
                print(f"{state_name_value}, {county_name_value}, {pop_change_percent}% Change in Population.")
        else:
            if pop_change_percent > 1.5:
                print(f"{state_name_value}, {county_name_value}, {pop_change_percent}% Change in Population.")


main()
