import pandas as pd
from docx import Document

file_path = r'C:\Users\Glori\PycharmProjects\_Project_School\Sofia_only_NVO.xlsx'
excel_data = pd.read_excel(file_path)
sofia_schools = excel_data.values.tolist()


def search_for_school(searched_school: str, list_of_schools: list):
    """search for a specific school in the entire list of schools"""
    is_found = False
    for row in list_of_schools:
        if row[1] == searched_school:
            is_found = True
            position_total = row[0]
            status_is_private, position_private_public = row[2].split(' / ')
            if status_is_private == 'Ч':
                is_private = 'private'
            else:
                is_private = 'public'
            position_private_public = int(position_private_public)
            count_students = int(row[5].split(' / ')[1])
            print(f'{searched_school} is a {is_private} school, ranked on {position_total} '
                  f'position among all schools in the capital.')
            if is_private == 'private':
                print(f'Among other private schools is ranked at {position_private_public} place.')
            else:
                print(f'Among other public schools is ranked at {position_private_public} place.')
            print(f'{count_students} pupils from the school attended the exam. ')

    if not is_found:
        print(f'I am sorry, but the school you are looking for is not in our database. '
              f'\nWould you like to see an alphabetical list of the schools?')
        second_choice = input('(Y/N)')
        if second_choice.lower() == 'n':
            main()
        elif second_choice.lower() == 'y':
            alphabetical_list = Document(
                r'C:\Users\Glori\PycharmProjects\_Project_School\alphabetical list of schools.docx')
            for row in alphabetical_list.paragraphs:
                print(row.text)
        else:
            print('Invalid answer.')
            main()


def compare_with_average(searched_school: str, list_of_schools: list):
    math_average_public = 72.64
    bel_average_sofia_public = 75.78
    math_average_private = 83.73
    bel_average_sofia_private = 84.75
    is_found = False
    for row in list_of_schools:
        if row[1] == searched_school:
            is_found = True
            math_2024 = float(row[5].split(' / ')[0])
            bel_2024 = float(row[4].split(' / ')[0])
            status_is_private = row[2][0]
            if status_is_private == 'Ч':
                if math_average_private > math_2024:
                    print(f'The results from the mathematics exam of {searched_school} school are lower than the '
                          f'average results for private schools.')
                elif math_average_private < math_2024:
                    print(f'The results from the mathematics exam of {searched_school} school are higher than the '
                          f'average results for private schools.')
                else:
                    print(f'The results from the mathematics exam of {searched_school} school are equal to the '
                          f'average results for private schools.')
                if bel_average_sofia_private > bel_2024:
                    print(f'The results from the BEL exam of {searched_school} school are lower than the '
                          f'average results for private schools.')
                elif bel_average_sofia_private < bel_2024:
                    print(f'The results from the BEL exam of {searched_school} school are higher than the '
                          f'average results for private schools.')
                else:
                    print(f'The results from the BEL exam of {searched_school} school are equal to the '
                          f'average results for private schools.')
            elif status_is_private == 'Д':
                if math_average_public > math_2024:
                    print(f'The results from the mathematics exam of {searched_school} school are lower than the '
                          f'average results for public schools.')
                elif math_average_public < math_2024:
                    print(f'The results from the mathematics exam of {searched_school} school are higher than the '
                          f'average results for public schools.')
                else:
                    print(f'The results from the mathematics exam of {searched_school} school are equal to the '
                          f'average results for public schools.')
                if bel_average_sofia_public > bel_2024:
                    print(f'The results from the BEL exam of {searched_school} school are lower than the '
                          f'average results for public schools.')
                elif bel_average_sofia_public < bel_2024:
                    print(f'The results from the BEL exam of {searched_school} school are higher than the '
                          f'average results for public schools.')
                else:
                    print(f'The results from the BEL exam of {searched_school} school are equal to the '
                          f'average results for public schools.')
    if not is_found:
        print(f'I am sorry, but the school you are looking for is not in our database.')
        main()


def compare_with_top(searched_school: str, list_of_schools: list):
    top_math_private = 96.62
    top_math_public = 90.09
    top_bel_private = 96.44
    top_bel_public = 91.29
    is_found = False
    for row in list_of_schools:
        if row[1] == searched_school:
            is_found = True
            math_2024 = float(row[5].split(' / ')[0])
            bel_2024 = float(row[4].split(' / ')[0])
            status_is_private = row[2][0]
            if status_is_private == 'Ч':
                if top_math_private > math_2024:
                    print(f'The best result among private schools in the mathematics exam is '
                          f'{(top_math_private - math_2024):.2f} higher than {searched_school} result.')
                else:
                    print(f'The school you are looking for has the top mathematics result - {top_math_private}!')
                if top_bel_private > bel_2024:
                    print(f'The best result among private schools in the BEL exam is '
                          f'{(top_bel_private - bel_2024):.2f} higher than {searched_school} result.')
                else:
                    print(f'The school you are looking for has the top BEL result - {top_bel_private}!')
            elif status_is_private == 'Д':
                if top_math_public > math_2024:
                    print(f'The best result among public schools in the mathematics exam is '
                          f'{(top_math_public - math_2024):.2f} higher than {searched_school} result.')
                else:
                    print(f'The school you are looking for has the top mathematics result - {top_math_public}!')
                if top_bel_public > bel_2024:
                    print(f'The best result among private schools in the BEL exam is '
                          f'{(top_bel_public - bel_2024):.2f} higher than {searched_school} result.')
                else:
                    print(f'The school you are looking for has the top BEL result - {top_bel_public}!')
    if not is_found:
        print(f'I am sorry, but the school you are looking for is not in our database.')
        main()


def compare_between_schools(first_school: str, second_school: str, list_of_schools: list):
    first_is_found = False
    second_is_found = False
    if first_school == second_school:
        print('You have to enter two different schools in order to be compared.')
        main()
    for row in list_of_schools:
        if row[1] == first_school:
            first_is_found = True
            first_math_2024 = float(row[5].split(' / ')[0])
            first_bel_2024 = float(row[4].split(' / ')[0])
        elif row[1] == second_school:
            second_is_found = True
            second_math_2024 = float(row[5].split(' / ')[0])
            second_bel_2024 = float(row[4].split(' / ')[0])
    if first_is_found and second_is_found:
        if first_math_2024 > second_math_2024 and first_bel_2024 > second_bel_2024:
            print(f'{first_school} has better results in both math and BEL exams.')
        elif first_math_2024 < second_math_2024 and first_bel_2024 < second_bel_2024:
            print(f'{second_school} has better results in both math and BEL exams.')
        elif first_math_2024 == second_math_2024 and first_bel_2024 == second_bel_2024:
            print(f'{first_school} and {second_school} schools have equal results in both math and BEL exams.')
        elif first_math_2024 > second_math_2024 and first_bel_2024 < second_bel_2024:
            print(f'{first_school} has better results in math and lower in the BEL exam.')
        elif first_math_2024 < second_math_2024 and first_bel_2024 > second_bel_2024:
            print(f'{first_school} has better results in BEL and lower in the math exam.')
        elif first_math_2024 == second_math_2024 and first_bel_2024 > second_bel_2024:
            print(f'{first_school} has better results in the BEL exam and both schools have equal math results.')
        elif first_math_2024 == second_math_2024 and first_bel_2024 < second_bel_2024:
            print(f'{second_school} has better results in the BEL exam and both schools have equal math results.')
        elif first_math_2024 > second_math_2024 and first_bel_2024 == second_bel_2024:
            print(f'{first_school} has better results in the math exam and both schools have equal BEL results.')
        elif first_math_2024 < second_math_2024 and first_bel_2024 == second_bel_2024:
            print(f'{second_school} has better results in the math exam and both schools have equal BEL results.')
    elif not first_is_found:
        print(f'I am sorry, but the first school you are looking for is not in our database.')
        main()
    elif not second_is_found:
        print(f'I am sorry, but the second school you are looking for is not in our database.')
        main()


def last_year_comparison(searched_school: str, list_of_schools: list):
    is_found = False
    for row in list_of_schools:
        if row[1] == searched_school:
            is_found = True
            math_2024 = float(row[5].split(' / ')[0])
            bel_2024 = float(row[4].split(' / ')[0])
            if row[7] is not None:
                try:
                    math_2023 = float(row[7].split(' / ')[0])
                except AttributeError:
                    math_2023 = None
                    print(f'We don\'t have information for {searched_school}\'s previous years results.')
                    main()
            if row[6] is not None:
                try:
                    bel_2023 = float(row[6].split(' / ')[0])
                except AttributeError:
                    bel_2023 = None
                    print(f'We don\'t have information for {searched_school}\'s previous years results.')
                    main()
            if row[9] is not None:
                try:
                    math_2022 = float(row[9].split(' / ')[0])
                except AttributeError:
                    math_2022 = None
                    print(f'We don\'t have information for {searched_school}\'s previous years results.')
                    main()
            if row[8] is not None:
                try:
                    bel_2022 = float(row[8].split(' / ')[0])
                except AttributeError:
                    bel_2022 = None
                    print(f'We don\'t have information for {searched_school}\'s previous years results.')
                    main()
            if math_2024 > math_2023 > math_2022:
                print(f'{searched_school} is improving its mathematics exam results since 2022.')
            elif math_2024 < math_2023 < math_2022:
                print(f'{searched_school}\' mathematics exam results are decreasing since 2022.')
            else:
                print(f'{searched_school}\' mathematics exam results vary in the years since 2022.')
            if bel_2024 > bel_2023 > bel_2022:
                print(f'{searched_school} is improving its BEL exam results since 2022.')
            elif bel_2024 > bel_2023 > bel_2022:
                print(f'{searched_school}\' BEL exam results are decreasing since 2022.')
            else:
                print(f'{searched_school}\' BEL exam results vary in the years since 2022.')
    if not is_found:
        print(f'I am sorry, but the school you are looking for is not in our database.')
        main()


def main():
    """ Main function to provide user interaction. """
    while True:
        print()
        print('***     School Comparison Assistant     ***')
        print('Here you can find simplified data based on '
              '\nschool exam results after 4th grade.')
        print('1 - School basic introduction')  # def search_for_school(searched_school: str, list_of_schools: list)
        print('2 - Compare with the average results')  # def compare_with_average(searched_school, list_of_schools)
        print('3 - Compare with the top result')  # def compare_with_top(searched_school: str, list_of_schools: list)
        print('4 - Comparison between two schools')  # def compare_between_schools(first, second, list_of_schools)
        print('5 - School progress in the years')  # def last_year_comparison(searched_school: str, list_of_schools)
        print('6 - Which is the best private school')  # top ranked private school info
        print('7 - Which is the best public school')  # top ranked public school info
        print('8 - Exit')
        print()

        choice = int(input('Enter your choice (1 - 8): '))

        if choice == 1:
            searched_school_name = input('Enter school name: ')
            search_for_school(searched_school_name, sofia_schools)
        elif choice == 2:
            searched_school_name = input('Enter school name: ')
            compare_with_average(searched_school_name, sofia_schools)
        elif choice == 3:
            searched_school_name = input('Enter school name: ')
            compare_with_top(searched_school_name, sofia_schools)
        elif choice == 4:
            first_school_name = input('Enter first school name: ')
            second_school_name = input('Enter second school name: ')
            compare_between_schools(first_school_name, second_school_name, sofia_schools)
        elif choice == 5:
            searched_school_name = input('Enter school name: ')
            last_year_comparison(searched_school_name, sofia_schools)
        elif choice == 6:
            print(f'Currently Private Primary School “St. Sofia” is the private school with highest results and over 50'
                  f' attendants. It is located in Gardova Glava, Vitosha district.')
            main()
        elif choice == 7:
            print(f'Currently 145 School “Simeon Radev” is the public school with highest results. '
                  f'It is located in Mladost 1a.')
            main()
        elif choice == 8:
            exit()


if __name__ == "__main__":
    main()
