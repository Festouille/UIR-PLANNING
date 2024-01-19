import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
import re
from icalendar import Calendar, Event

def extract_schedule_data(driver):
    js_script = """
    var elements = Array.from(document.querySelectorAll('.d-flex.align-items-center.mb-6'));
    var data = [];
    var joursSemaine = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];

    for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        var text = element.innerText.trim();

        if (!text.startsWith("Pas de planning disponible")) {
            data.push(joursSemaine[i % joursSemaine.length] + "\\n" + text);
        }
    }

    return data;
    """

    data = driver.execute_script(js_script)
    return data

def extract_start_date(week_text):
    match = re.search(r'\d{2}-\d{2}-\d{4}', week_text)
    return match.group(0) if match else None

def create_icalendar(df):
    cal = Calendar()
    for _, row in df.iterrows():
        event = Event()
        event.add('summary', row['Cours'])
        start_date = pd.to_datetime(row['Semaine'], format='%d-%m-%Y')
        event.add('dtstart', start_date)
        event.add('dtend', start_date + pd.Timedelta(hours=3))
        cal.add_component(event)
    return cal

def save_icalendar(cal, filepath):
    with open(filepath, 'wb') as f:
        f.write(cal.to_ical())

def display_calendar(df):
    pd.set_option('display.max_colwidth', None)
    print(df.to_string(index=False))

chromedriver_path = "C://Users//akram//OneDrive//Documents//chromedriver-win64//chromedriver.exe"
url = "https://eservices.uir.ac.ma/planning"
username = os.getenv("UIR_USERNAME")
password = os.getenv("UIR_PASSWORD")

if not username or not password:
    raise ValueError("Veuillez définir les variables d'environnement UIR_USERNAME et UIR_PASSWORD.")

try:
    with webdriver.Chrome() as driver:
        driver.get(url)

        email_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "email"))
        )
        email_field.send_keys(username)

        password_field = driver.find_element(By.NAME, "password")
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".d-flex.align-items-center.mb-6"))
        )

        all_week_data = []

        for _ in range(4):
            try:
                driver.find_element(By.XPATH, '//*[@id="kt_app_content"]/div/div/div/div/div[1]/div[3]/button').click()
                time.sleep(5)

                data = extract_schedule_data(driver)
                current_week_date_range = driver.find_element(By.XPATH, '//*[@id="kt_app_content"]/div/div/div/div/div[1]/div[2]').text.strip()
                all_week_data.append((current_week_date_range, data))

            except WebDriverException as e:
                print(f"Une exception WebDriver s'est produite : {str(e)}")
            except Exception as e:
                print(f"Une exception s'est produite : {str(e)}")

finally:
    driver.quit()

df = pd.DataFrame(all_week_data, columns=["Semaine", "Données"])
df["Start_Date"] = df["Semaine"].apply(extract_start_date)

courses = [(row["Start_Date"], course) for _, row in df.iterrows() for course in row["Données"] if not course.startswith("Pas de planning disponible")]
new_df = pd.DataFrame(courses, columns=["Semaine", "Cours"])
new_df['Cours'] = new_df['Cours'].str.replace('\n', ' ')
new_df.reset_index(drop=True, inplace=True)

pd.set_option('display.max_colwidth', None)
display_calendar(new_df)

save_to_ical = input("Voulez-vous sauvegarder le calendrier au format .ics? (Oui/Non): ").lower() == "oui"

if save_to_ical:
    icalendar_path = 'emploi_du_temps.ics'
    cal = create_icalendar(new_df)
    save_icalendar(cal, icalendar_path)
    print(f"Calendrier sauvegardé au format .ics : {os.path.abspath(icalendar_path)}")
