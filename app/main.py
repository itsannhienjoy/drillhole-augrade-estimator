import processing


if __name__ == "__main__":
    file_name = "app/TestData.xlsx"
    new_hole = (2.5, 32.5)

    proc = processing.ProcessPartA()
    proc.retrieve_data(file_name)

    # Task 1 
    avg_au_grade = proc.calculate_average_au_grade()
    print(f"\n✅ Task 1: Average Au Grade = {avg_au_grade:.2f}")

    # Task 6
    company_totals = proc.get_total_drilled_by_company()
    print("\n✅ Task 6: Total Drilled Meters by Company:")
    for company, meters in company_totals.items():
        print(f"\n{company}: {meters} meters")

    # Task 7
    total_length =  proc.get_daily_drilled_meters()
    print("\n✅ Task 7: Drilled Meters by Day:")
    for day, meters in total_length.items():
        print(f"\n{day}: {meters} meters")

    # Task 8 & 9
    proc.print_distances_to_point(*new_hole)
    proc.estimate_augrade_from_nearest(*new_hole)
   
    # Task 10
    new_holes = [
        (7.5, 32.5),  
        (2.5, 37.5),  
        (7.5, 37.5)  
    ]
    print("\n✅ Task 10: Estimated Au Grades For New Drillholes:")
    for (x, y) in new_holes:
        proc.estimate_augrade_from_nearest(x, y)

    print('Done')
