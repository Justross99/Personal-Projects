from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import os
import openpyxl
import datetime
from screeninfo import get_monitors


def workoutentry():
    #spotify = subprocess.Popen(['C:/your/path/here/Spotify.exe'])
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option(
        "excludeSwitches", ['enable-automation'])
    chrome_options.add_argument("start-maximized")
    browser1 = webdriver.Chrome(
        ChromeDriverManager().install(), options=chrome_options)
    browser1.get('https://www.google.com/search?q=stopwatch')
    browser1.execute_script("window.open('https://open.spotify.com/')")

    for m in get_monitors():
        if m.is_primary:
            monitor_width = m.width
            monitor_height = m.height
            break

    browser1.set_window_size(monitor_width//2+10, monitor_height-40)
    browser1.set_window_position(monitor_width//2, 0)
    browser2 = webdriver.Chrome(
        ChromeDriverManager().install(), options=chrome_options)
    browser2.get(
        "https://darebee.com/pdf/programs/30-days-of-hiit-advanced.pdf")
    browser2.set_window_position(0, 0)
    browser2.set_window_size(monitor_width//2+50, monitor_height-40)

    date_last_done = Workout_entries["A2"].value
    new_date_last_done = date_last_done + 1

    actions = ActionChains(browser2)
    for _ in range(new_date_last_done):
        actions.send_keys(Keys.ARROW_RIGHT).perform()

    columnC = []
    cellcheckcounter = -1
    colC = Workout_entries["C"]
    for day in colC:
        if day.value is None:
            break
        columnC.append(day.value)
        cellcheckcounter += 1
    bottomrow = cellcheckcounter + 2

    date = datetime.date.today()
    print("today is day: " + str(new_date_last_done))
    completed = input("Did you finish your workout today?")
    if completed == "yes" or completed == "Yes":
        Workout_entries["C" + str(bottomrow)].value = date
        if new_date_last_done >= 31:
            new_date_last_done = 0
        Workout_entries["A2"].value = new_date_last_done
        if new_date_last_done == 0:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(
                bottomrow)].value = "30 sec high knees, 30 sec push ups, 30 sec punches, 30 sec high knees, 30 sec plank walk-outs, 30 sec punches, 2 min rest"
        if new_date_last_done == 1:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(
                bottomrow)].value = "20 sec high knees, 20 sec punches, 20 sec plank+jab+cross, 20 sec high knees, 20 sec punches, 20 sec plank+jab+cross, 2 min rest"
        if new_date_last_done == 2:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(
                bottomrow)].value = "40 sec squats, 20 sec shoulder taps, 40 sec side kicks, 20 sec push ups, 2 min rest"
        if new_date_last_done == 3:
            Workout_entries["D" + str(bottomrow)].value = 10
            Workout_entries["E" + str(
                bottomrow)].value = "20 sec high knees, 10 sec march in place, 20 sec climbers, 10 sec slow climbers"
        if new_date_last_done == 4:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(
                bottomrow)].value = "20 sec squats, 20 sec plank walk-outs, 20 sec scissor chops, 20 sec squats, 20 sec push ups, 20 sec arm scissors, 2 min rest"
        if new_date_last_done == 5:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)
                            ].value = "4x(20 sec burpees, 10 sec rest), 2 min rest"
        if new_date_last_done == 6:
            Workout_entries["D" + str(bottomrow)].value = 10
            Workout_entries["E" + str(
                bottomrow)].value = "20 sec jumping jacks, 10 sec step jacks, 20 sec high knees, 10 sec march in place"
        if new_date_last_done == 7:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(
                bottomrow)].value = "4x(20 sec punches, 10 sec push up+shoulder tap), 2 min rest"
        if new_date_last_done == 8:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)].value = "20 sec high knees, 20 sec climbers, 20 sec high knees, 20 sec plank jacks, 20 sec high knees, 20 sec plank jump-ins, push up every 20 sec, 2 min rest"
        if new_date_last_done == 9:
            Workout_entries["D" + str(bottomrow)].value = 10
            Workout_entries["E" + str(
                bottomrow)].value = "4x(20 sec jump squats, 10 sec step jacks, 20 sec push ups, 10 sec punches)"
        if new_date_last_done == 10:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)
                            ].value = "4x(20 sec high knees, 10 sec plank), 2 min rest"
        if new_date_last_done == 11:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)].value = "20 sec jumping jacks, 20 sec arm circles, 20 sec jumping jacks, 20 sec scissor chops, 20 sec jumping jacks, 20 sec arm scissors, side-to-side lunge every 20 sec, 2 min rest"
        if new_date_last_done == 12:
            Workout_entries["D" + str(bottomrow)].value = 30
            Workout_entries["E" + str(
                bottomrow)].value = "3 supersets of punches, squats, push ups, climbers, punches, jump squats, shoulder taps, burpees"
        if new_date_last_done == 13:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)].value = "20 sec high knees, 20 sec knee-to-elbow, 20 sec high knees, 20 sec burpees, 20 sec high knees, 20 sec plank jack burpees, one push up every 20 sec, 2 min rest"
        if new_date_last_done == 14:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)
                            ].value = "4x(20 sec squats, 10 sec push ups), 2 min rest"
        if new_date_last_done == 15:
            Workout_entries["D" + str(bottomrow)].value = 20
            Workout_entries["E" + str(bottomrow)
                            ].value = "40x(20 sec high knees, 10 sec rest)"
        if new_date_last_done == 16:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(
                bottomrow)].value = "30 sec high knees, 10 sec shoulder taps, 20 sec punches, 30 sec high knees, 10 sec plank walk-outs, 20 sec punches, 2 min rest"
        if new_date_last_done == 17:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(
                bottomrow)].value = "40 sec side kicks, 20 sec push ups, 40 sec lunges, 20 sec push ups, 2 min rest"
        if new_date_last_done == 18:
            Workout_entries["D" + str(bottomrow)].value = 20
            Workout_entries["E" + str(
                bottomrow)].value = "30 sec jumping jacks, 30 sec step jacks, 30 sec high knees, 30 sec march in place"
        if new_date_last_done == 19:
            Workout_entries["D" + str(bottomrow)].value = 40
            Workout_entries["E" + str(bottomrow)
                            ].value = "4x(20 sec burpees, 10 sec rest), 2 min rest"
        if new_date_last_done == 20:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(bottomrow)
                            ].value = "3x(30 sec high knees, 30 sec plank hold), 2 min rest"
        if new_date_last_done == 21:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(
                bottomrow)].value = "30 sec squat, 30 sec plank walk-outs, 30 sec scissor chops, 30 sec squats, 30 sec push ups, 30 sec arm scissors, 2 min rest"
        if new_date_last_done == 22:
            Workout_entries["D" + str(bottomrow)].value = 20
            Workout_entries["E" + str(
                bottomrow)].value = "30 sec high knees, 30 sec march in place, 30 sec climbers, 30 sec slow climbers"
        if new_date_last_done == 23:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(
                bottomrow)].value = "3x(30 sec punches, 30 sec push up + shoulder tap), 2 min rest"
        if new_date_last_done == 24:
            Workout_entries["D" + str(bottomrow)].value = 30
            Workout_entries["E" + str(
                bottomrow)].value = "3 supersets of punches, squats, arm scissors, climbers, punches, jump squats, scissor chops, burpees"
        if new_date_last_done == 25:
            Workout_entries["D" + str(bottomrow)].value = 30
            Workout_entries["E" + str(bottomrow)
                            ].value = "10x(20 sec high knees, 10 sec rest), 2 min rest"
        if new_date_last_done == 26:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(bottomrow)].value = "30 sec jumping jacks, 30 sec arm circles, 30 sec jumping jacks, 30 sec scissor chops, 30 sec jumping jacks, 30 sec arm scissors, one side-to-side lunge every 30 sec, 2 min rest"
        if new_date_last_done == 27:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(bottomrow)
                            ].value = "3x(30 sec squats, 30 sec plank), 2 min rest"
        if new_date_last_done == 28:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(bottomrow)].value = "30 sec high knees, 30 sec punches, 30 sec plank+jab+cross, 30 sec high knees, 30 sec punches, 30 sec plank jack+jab+cross, one push up every 30 sec, 2 min rest"
        if new_date_last_done == 29:
            Workout_entries["D" + str(bottomrow)].value = 50
            Workout_entries["E" + str(bottomrow)
                            ].value = "4x(30 sec burpees, 10 sec rest), 2 min rest"
        Workout_entries["A2"].value = new_date_last_done

        wb.save("./Nutrition.xlsx")
        browser1.quit()
        browser2.quit()

    else:
        browser1.quit()
        browser2.quit()


wb = openpyxl.load_workbook(os.path.join(
    os.path.dirname(__file__), "./Nutrition.xlsx"))
sheetlist = wb.sheetnames
Workout_entries = wb["Workout entries"]

workoutentry()
