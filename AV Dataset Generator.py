import math

import cv2
import easyocr
import numpy as np
from pdf2image import convert_from_path
import openpyxl

from tkinter.filedialog import askopenfilenames

filenames = askopenfilenames(title="Select file", filetype=[("Pdf file", "*.pdf")])

wb = openpyxl.Workbook()
ws = wb.active

headers = (
    "Manufacturer Name", "AVT Number", "Business Name", "Telephone Number", "Street Address", "City", "State",
    "Zip Code",
    "Date", "Time", "Vehicle 1 Year", "Vehicle 1 make", "Vehicle 1 model", "Vehicle 1 License Plate Number",
    "Vehicle 1 Identification Number",
    "Vehicle 1 Registered In", "Address", "City", "Country", "State", "Zip Code", "Vehicle 1 Status",
    "Vehicles Involved (Section 2)",
    "Vehicle 1 Driver Full Name", "Vehicle 1 Driver License Number", "State", "Date of Birth", "Insurance Company Name",
    "Policy Number", "Company NAIC Number",
    "Policy Period", "Vehicle 1 Damage", "Vehicle 2 Year", "Vehicle 2 Model", "Vehicle 2 License Plate Number",
    "Vehicle 2 Identification Number",
    "Vehicle 2 Registered In", "Vehicle 2 Status", "Vehicles Involved (Section 3)", "Vehicle 2 Driver Full Name",
    "Vehicle 2 Driver License Number", "State", "Date of Birth",
    "Insurance Company Name", "Policy Number", "Company NAIC Number", "Policy Period", "Mode", "Description",
    "Vehicle 1 Weather", "Vehicle 1 Weather other Choice", "Vehicle 2 Weather", "Vehicle 2 Weather other Choice",
    "Vehicle 1 Lighting", "Vehicle 1 Lighting other Choice", "Vehicle 2 Lighting", "Vehicle 2 Lighting other Choice",
    "Vehicle 1 Roadway Surface", "Vehicle 1 Roadway Surface other Choice", "Vehicle 2 Roadway Surface",
    "Vehicle 2 Roadway Surface other Choice",
    "Vehicle 1 Roadway Condition", "Vehicle 1 Roadway Condition other Choice", "Vehicle 2 Roadway Condition",
    "Vehicle 2 Roadway Condition other Choice",
    "Vehicle 1 Movement", "Vehicle 1 Movement other Choice", "Vehicle 2 Movement", "Vehicle 2 Movement other Choice",
    "Vehicle 1 Collision", "Vehicle 1 Collision other Choice", "Vehicle 2 Collision",
    "Vehicle 2 Collision other Choice")
ws.append(headers);

from datetime import datetime

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)

for idx, i in enumerate(filenames):
    pageOneExist = False
    pageTwoExist = False
    pageThreeExist = False
    print(f"Processing File Number {idx}")
    images = convert_from_path(i,
                               poppler_path=r"C:\Users\SyntaxError\Desktop\Release-23.01.0-0\poppler-23.01.0\Library\\bin")

    if len(images) > 0:
        pageOne = np.array(images[0].convert('RGB'))
        pageOneExist = True
    if len(images) > 1:
        pageTwo = np.array(images[1].convert('RGB'))
        pageTwoExist = True
    if len(images) > 2:
        pageThreeExist = True
        pageThree = np.array(images[2].convert('RGB'))

    s1_manu_name = ""
    s1_avt = ""
    s1_business_name = ""
    s1_telephone_number = ""
    s1_street_address = ""
    s1_city = ""
    s1_state = ""
    s1_zip = ""
    s2_date = ""
    s2_time = ""
    s2_vehicle_year = ""
    s2_make = ""
    s2_model = ""
    s2_license_plate_number = ""
    s2_vehicle_identification_number = ""
    s2_sviri = ""
    s2_address = ""
    s2_city = ""
    s2_country = ""
    s2_state = ""
    s2_zip = ""
    vehicle_one_status = ""
    s2_vehicles_involved = ""
    s2_driver_fullname = ""
    s2_driver_license_number = ""
    s2_driver_state = ""
    s2_driver_dob = ""
    s2_driver_insurance = ""
    s2_driver_policy_number = ""
    s2_naic_number = ""
    s2_policy_period = ""
    vehicle_one_damage = ""
    s3_vehicle_year = ""
    s3_model = ""
    s3_license_plate_number = ""
    s3_vehicle_identification_number = ""
    vehicle_two_status = ""
    s3_sviri = ""
    s3_vehicles_involved = ""
    s3_driver_fullname = ""
    s3_driver_license = ""
    s3_state = ""
    s3_driver_dob = ""
    s3_driver_insurance = ""
    s3_driver_policy_number = ""
    s3_naic_number = ""
    s3_policy_period = ""
    mode = ""
    description = ""
    vehicle_one_weather = ""
    vehicle_one_other_weather_choices = ""
    vehicle_two_weather = ""
    vehicle_two_other_weather_choices = ""
    vehicle_one_lighting = ""
    vehicle_one_other_lighting_choices = ""
    vehicle_two_lighting = ""
    vehicle_two_other_lighting_choices = ""
    vehicle_one_roadway_surface = ""
    vehicle_one_other_roadway_surface_choices = ""
    vehicle_two_roadway_surface = ""
    vehicle_two_other_roadway_surface_choices = ""
    vehicle_one_roadway_condition = ""
    vehicle_one_other_roadway_condition_choices = ""
    vehicle_two_roadway_condition = ""
    vehicle_two_other_roadway_condition_choices = ""
    vehicle_one_movement = ""
    vehicle_one_other_movement_choices = ""
    vehicle_two_movement = ""
    vehicle_two_other_movement_choices = ""
    vehicle_one_collision_type = ""
    vehicle_one_other_collision_type = ""
    vehicle_two_collision_type = ""
    vehicle_two_other_collision_type = ""

    reader = easyocr.Reader(['en'])


    def detect_checkbox(img):
        gray_scale = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        th1, img_bin = cv2.threshold(gray_scale, 150, 225, cv2.THRESH_BINARY)
        img_bin = ~ img_bin
        line_min_width = 20
        kernal_h = np.ones((1, line_min_width), np.uint8)
        kernal_v = np.ones((line_min_width, 1), np.uint8)
        img_bin_h = cv2.morphologyEx(img_bin, cv2.MORPH_OPEN, kernal_h)
        img_bin_v = cv2.morphologyEx(img_bin, cv2.MORPH_OPEN, kernal_v)
        img_bin_final = img_bin_h | img_bin_v
        final_kernel = np.ones((3, 3), np.uint8)
        img_bin_final = cv2.dilate(img_bin_final, final_kernel, iterations=1)
        ret, labels, stats, centroids = cv2.connectedComponentsWithStats(~img_bin_final, connectivity=8,
                                                                         ltype=cv2.CV_32S)
        # cv2.imshow('', img_bin_final)
        # cv2.waitKey(0)
        return stats


    # Section 1

    # Manufacturer's Name
    s1_manu_name_img = pageOne[960:990, 65:1270]
    s1_manu_name = reader.readtext(s1_manu_name_img, detail=0)
    s1_manu_name = ''.join(s1_manu_name)

    # AVT Number
    s1_avt_img = pageOne[950:995, 1288:1628]
    s1_avt = reader.readtext(s1_avt_img, detail=0)
    s1_avt = ''.join(s1_avt)

    # Business Name
    s1_business_name_img = pageOne[1025:1060, 65:1275]
    s1_business_name = reader.readtext(s1_business_name_img, detail=0)
    s1_business_name = ''.join(s1_business_name)

    # Telephone Number
    s1_telephone_number_img = pageOne[1017:1060, 1288:1629]
    s1_telephone_number = reader.readtext(s1_telephone_number_img, detail=0)
    s1_telephone_number = ''.join(s1_telephone_number)

    # Street Address
    s1_street_address_img = pageOne[1085:1125, 65:660]
    s1_street_address = reader.readtext(s1_street_address_img, detail=0)
    s1_street_address = ''.join(s1_street_address)

    # City
    s1_city_img = pageOne[1085:1125, 665:1280]
    s1_city = reader.readtext(s1_city_img, detail=0)
    s1_city = ''.join(s1_city)

    # State
    s1_state_img = pageOne[1085:1125, 1285:1390]
    s1_state = reader.readtext(s1_state_img, detail=0)
    s1_state = ''.join(s1_state)

    # Zip Code
    s1_zip_img = pageOne[1085:1125, 1400:1635]
    s1_zip = reader.readtext(s1_zip_img, detail=0)
    s1_zip = ''.join(s1_zip)

    # Section 2

    # Date
    s2_date_img = pageOne[1225:1265, 65:350]
    s2_date = reader.readtext(s2_date_img, detail=0)
    s2_date = ''.join(s2_date)

    # Time
    s2_time_img = pageOne[1225:1265, 370:470]
    s2_time = reader.readtext(s2_time_img, detail=0)
    s2_time = ''.join(s2_time)
    am_pm_image = pageOne[1230:1266, 475:663]
    grayImage = cv2.cvtColor(am_pm_image, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage

    stats = detect_checkbox(am_pm_image)
    time_list = ['AM', 'PM']
    sum_of_sums = []
    for x, y, w, h, area in stats[2:]:
        checkbox = blackAndWhiteImage[y:y + h, x: x + w]
        sum_of_sums.append(np.sum(checkbox))
    if len(sum_of_sums) > 0:
        if len(time_list) > np.argmax(sum_of_sums):
            s2_time += ' ' + time_list[np.argmax(sum_of_sums)]

    # Vehicle Year
    s2_vehicle_year_img = pageOne[1220:1260, 670:965]
    s2_vehicle_year = reader.readtext(s2_vehicle_year_img, detail=0)
    s2_vehicle_year = ''.join(s2_vehicle_year)

    # Make
    s2_make_img = pageOne[1225:1265, 970:1280]
    s2_make = reader.readtext(s2_make_img, detail=0)
    s2_make = ''.join(s2_make)

    # Model
    s2_model_img = pageOne[1225:1265, 1290:1625]
    s2_model = reader.readtext(s2_model_img, detail=0)
    s2_model = ''.join(s2_model)

    # License Plate Number
    s2_license_plate_number_img = pageOne[1285:1330, 60:365]
    s2_license_plate_number = reader.readtext(s2_license_plate_number_img, detail=0)
    s2_license_plate_number = ''.join(s2_license_plate_number)

    # Vehicle Identification Number
    s2_vehicle_identification_number_img = pageOne[1285:1330, 368:1284]
    s2_vehicle_identification_number = reader.readtext(s2_vehicle_identification_number_img, detail=0)
    s2_vehicle_identification_number = ''.join(s2_vehicle_identification_number)

    # State Vehicle is Registered in
    s2_sviri_img = pageOne[1285:1330, 1290:1635]
    s2_sviri = reader.readtext(s2_sviri_img, detail=0)
    s2_sviri = ''.join(s2_sviri)

    # Address
    s2_address_img = pageOne[1360:1395, 60:650]
    s2_address = reader.readtext(s2_address_img, detail=0)
    s2_address = ''.join(s2_address)

    # City
    s2_city_img = pageOne[1360:1395, 660:965]
    s2_city = reader.readtext(s2_city_img, detail=0)
    s2_city = ''.join(s2_city)

    # Country
    s2_country_img = pageOne[1360:1395, 1035:1285]
    s2_country = reader.readtext(s2_country_img, detail=0)
    s2_country = ''.join(s2_country)

    # State
    s2_state_img = pageOne[1360:1395, 1285:1400]
    s2_state = reader.readtext(s2_state_img, detail=0)
    s2_state = ''.join(s2_state)

    # Zip Code
    s2_zip_img = pageOne[1360:1395, 1400:1635]
    s2_zip = reader.readtext(s2_zip_img, detail=0)
    s2_zip = ''.join(s2_zip)

    # Vehicle One Status
    vehicle_one_status_img = pageOne[1407:1467, 240:510]
    grayImage = cv2.cvtColor(vehicle_one_status_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    stats = detect_checkbox(vehicle_one_status_img)
    checkbox_list = ['Moving', 'Stopped in Traffic']
    sum_of_sums = []
    for x, y, w, h, area in stats[2:]:
        checkbox = blackAndWhiteImage[y:y + h, x: x + w]
        sum_of_sums.append(np.sum(checkbox))
    if len(sum_of_sums) > 0:
        if len(checkbox_list) > np.argmax(sum_of_sums):
            vehicle_one_status = checkbox_list[np.argmax(sum_of_sums)]

    # Number Of Vehicles Involved
    s2_vehicles_involved_img = pageOne[1426:1464, 1291:1624]
    grayImage = cv2.cvtColor(s2_vehicles_involved_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    final_kernel = np.ones((3, 3), np.uint8)
    blackAndWhiteImage = cv2.dilate(blackAndWhiteImage, final_kernel, iterations=1)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    s2_vehicles_involved = reader.readtext(blackAndWhiteImage, detail=0)
    if len(s2_vehicles_involved) > 0:
        s2_vehicles_involved = ''.join(s2_vehicles_involved[0])
    else:
        s2_vehicles_involved = ""

    # Driver's Full Name
    s2_driver_fullname_img = pageOne[1490:1530, 60:745]
    s2_driver_fullname = reader.readtext(s2_driver_fullname_img, detail=0)
    s2_driver_fullname = ''.join(s2_driver_fullname)

    # Driver's License Number
    s2_driver_license_number_img = pageOne[1490:1530, 750:1284]
    s2_driver_license_number = reader.readtext(s2_driver_license_number_img, detail=0)
    s2_driver_license_number = ''.join(s2_driver_license_number)

    # Driver's State
    s2_driver_state_img = pageOne[1490:1530, 1287:1396]
    s2_driver_state = reader.readtext(s2_driver_state_img, detail=0)
    s2_driver_state = ''.join(s2_driver_state)

    # Driver's DOB
    s2_driver_dob_img = pageOne[1490:1530, 1400:1635]
    s2_driver_dob = reader.readtext(s2_driver_dob_img, detail=0)
    s2_driver_dob = ''.join(s2_driver_dob)

    # Driver's insurance
    s2_driver_insurance_img = pageOne[1555:1595, 60:747]
    s2_driver_insurance = reader.readtext(s2_driver_insurance_img, detail=0)
    s2_driver_insurance = ''.join(s2_driver_insurance)

    # Driver's Policy Number
    s2_driver_policy_number_img = pageOne[1555:1595, 750:1635]
    s2_driver_policy_number = reader.readtext(s2_driver_policy_number_img, detail=0)
    s2_driver_policy_number = ''.join(s2_driver_policy_number)

    # Company NAIC Number
    s2_naic_number_img = pageOne[1620:1660, 60:745]
    s2_naic_number = reader.readtext(s2_naic_number_img, detail=0)
    s2_naic_number = ''.join(s2_naic_number)

    # Policy Period
    s2_policy_period_from_img = pageOne[1620:1652, 830:1205]
    s2_policy_period_from = reader.readtext(s2_policy_period_from_img, detail=0)
    s2_policy_period_from = ''.join(s2_policy_period_from)
    s2_policy_period_to_img = pageOne[1620:1652, 1255:1625]
    s2_policy_period_to = reader.readtext(s2_policy_period_to_img, detail=0)
    s2_policy_period_to = ''.join(s2_policy_period_to)
    s2_policy_period = s2_policy_period_from + "-" + s2_policy_period_to

    # Damage
    cv2.rectangle(pageOne, (185, 1725), (735, 1865), (255, 0, 0), 2)
    damage_checkbox = pageOne[1725:1865, 185:735]
    grayImage = cv2.cvtColor(damage_checkbox, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage

    stats = detect_checkbox(damage_checkbox)
    checkbox_list = ['UNK', 'NONE', 'MINOR', 'MOD', 'MAJOR']
    sum_of_sums = []
    for x, y, w, h, area in stats[2:]:
        checkbox = blackAndWhiteImage[y:y + h, x: x + w]
        sum_of_sums.append(np.sum(checkbox))
    if len(sum_of_sums) > 0:
        if len(checkbox_list) > np.argmax(sum_of_sums):
            vehicle_one_damage = checkbox_list[np.argmax(sum_of_sums)]

    #     # PAGE TWO STARTED
    if pageTwoExist is not True:
        output = (
            s1_manu_name, s1_avt, s1_business_name, s1_telephone_number, s1_street_address,
            s1_city, s1_state, s1_zip, s2_date, s2_time, s2_vehicle_year, s2_make, s2_model,
            s2_license_plate_number, s2_vehicle_identification_number, s2_sviri, s2_address,
            s2_city, s2_country, s2_state, s2_zip, vehicle_one_status, s2_vehicles_involved,
            s2_driver_fullname, s2_driver_license_number, s2_driver_state, s2_driver_dob,
            s2_driver_insurance, s2_driver_policy_number, s2_naic_number, s2_policy_period,
            vehicle_one_damage, s3_vehicle_year, s3_model, s3_license_plate_number,
            s3_vehicle_identification_number, vehicle_two_status, s3_sviri, s3_vehicles_involved, s3_driver_fullname,
            s3_driver_license, s3_state, s3_driver_dob, s3_driver_insurance, s3_driver_policy_number,
            s3_naic_number, s3_policy_period, mode, ' '.join(description), vehicle_one_weather,
            vehicle_one_other_weather_choices, vehicle_two_weather, vehicle_two_other_weather_choices,
            vehicle_one_lighting, vehicle_one_other_lighting_choices, vehicle_two_lighting,
            vehicle_two_other_lighting_choices, vehicle_one_roadway_surface,
            vehicle_one_other_roadway_surface_choices, vehicle_two_roadway_surface,
            vehicle_two_other_roadway_surface_choices, vehicle_one_roadway_condition,
            vehicle_one_other_roadway_condition_choices, vehicle_two_roadway_condition,
            vehicle_two_other_roadway_condition_choices, vehicle_one_movement,
            vehicle_one_other_movement_choices, vehicle_two_movement,
            vehicle_two_other_movement_choices, vehicle_one_collision_type,
            vehicle_one_other_collision_type, vehicle_two_collision_type,
            vehicle_two_other_collision_type)

        ws.append(output)
        continue

    # SECTION 3
    # Vehicle Year
    s3_vehicle_year_img = pageTwo[150:197, 60:366]
    s3_vehicle_year = reader.readtext(s3_vehicle_year_img, detail=0)
    s3_vehicle_year = ''.join(s3_vehicle_year)

    # Vehicle Model
    s3_model_img = pageTwo[150:197, 368:1632]
    s3_model = reader.readtext(s3_model_img, detail=0)
    s3_model = ''.join(s3_model)

    # License Plate
    s3_license_plate_number_img = pageTwo[218:264, 60:366]
    s3_license_plate_number = reader.readtext(s3_license_plate_number_img, detail=0)
    s3_license_plate_number = ''.join(s3_license_plate_number)

    # Vehicle Identification Number
    s3_vehicle_identification_number_img = pageTwo[218:264, 368:1284]
    s3_vehicle_identification_number = reader.readtext(s3_vehicle_identification_number_img, detail=0)
    s3_vehicle_identification_number = ''.join(s3_vehicle_identification_number)

    # State Vehicle Registered in

    s3_sviri_img = pageTwo[218:264, 1287:1630]
    s3_sviri = reader.readtext(s3_sviri_img, detail=0)
    s3_sviri = ''.join(s3_sviri)

    # Vehicles Involved

    s3_vehicles_involved_img = pageTwo[284:330, 1287:1630]
    s3_vehicles_involved = reader.readtext(s3_vehicles_involved_img, detail=0)
    s3_vehicles_involved = ''.join(s3_vehicles_involved)

    # driver fullname
    s3_driver_fullname_img = pageTwo[355:397, 65:747]
    s3_driver_fullname = reader.readtext(s3_driver_fullname_img, detail=0)
    s3_driver_fullname = ''.join(s3_driver_fullname)

    # driver license
    s3_driver_license_img = pageTwo[355:397, 750:1284]
    s3_driver_license = reader.readtext(s3_driver_license_img, detail=0)
    s3_driver_license = ''.join(s3_driver_fullname)

    # State
    s3_state_img = pageTwo[355:397, 1287:1396]
    s3_state = reader.readtext(s3_state_img, detail=0)
    s3_state = ''.join(s3_state)

    # DOB
    s3_driver_dob_img = pageTwo[355:397, 1398:1630]
    s3_driver_dob = reader.readtext(s3_driver_dob_img, detail=0)
    s3_driver_dob = ''.join(s3_driver_dob)

    # Insurance Company Name
    s3_driver_insurance_img = pageTwo[420:464, 65:747]
    s3_driver_insurance = reader.readtext(s3_driver_insurance_img, detail=0)
    s3_driver_insurance = ''.join(s3_driver_insurance)

    # Policy Number
    s3_driver_policy_number_img = pageTwo[420:464, 750:1630]
    s3_driver_policy_number = reader.readtext(s3_driver_policy_number_img, detail=0)
    s3_driver_policy_number = ''.join(s3_driver_policy_number)

    # NAIC Number
    s3_naic_number_img = pageTwo[485:530, 65:747]
    s3_naic_number = reader.readtext(s3_naic_number_img, detail=0)
    s3_naic_number = ''.join(s3_naic_number)

    # Policy Period
    s3_policy_period_from_img = pageTwo[485:519, 835:1205]
    s3_policy_period_from = reader.readtext(s3_policy_period_from_img, detail=0)
    s3_policy_period_from = ''.join(s3_policy_period_from)
    s3_policy_period_to_img = pageTwo[485:519, 1260:1630]
    s3_policy_period_to = reader.readtext(s3_policy_period_to_img, detail=0)
    s3_policy_period_to = ''.join(s3_policy_period_to)
    s3_policy_period = s3_policy_period_from + "-" + s3_policy_period_to

    # Section 4

    # Name
    s4_name1_img = pageTwo[680:722, 65:1635]
    s4_name1 = reader.readtext(s4_name1_img, detail=0)
    s4_name1 = ''.join(s4_name1)

    # Address
    s4_address1_img = pageTwo[743:789, 65:670]
    s4_address1 = reader.readtext(s4_address1_img, detail=0)
    s4_address1 = ''.join(s4_address1)

    # City
    s4_city1_img = pageTwo[743:789, 670:1285]
    s4_city1 = reader.readtext(s4_city1_img, detail=0)
    s4_city1 = ''.join(s4_city1)

    # State
    s4_state1_img = pageTwo[743:789, 1290:1400]
    s4_state1 = reader.readtext(s4_state1_img, detail=0)
    s4_state1 = ''.join(s4_state1)

    # Zip Code
    s4_zip_img = pageTwo[743:789, 1400:1635]
    s4_zip = reader.readtext(s4_zip_img, detail=0)
    s4_zip = ''.join(s4_zip)

    # Description
    cv2.rectangle(pageTwo, (60, 1695), (1635, 2055), (255, 0, 0), 2)
    description_img = pageTwo[1695:2055, 60:1635]
    description = reader.readtext(description_img, detail=0)

    # Vehicle Two Status
    cv2.rectangle(pageTwo, (240, 268), (505, 329), (255, 0, 0), 2)
    vehicle_two_status_img = pageTwo[268:329, 240:505]
    grayImage = cv2.cvtColor(vehicle_two_status_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    stats = detect_checkbox(vehicle_two_status_img)
    checkbox_list = ['Moving', 'Stopped in Traffic']
    sum_of_sums = []
    for x, y, w, h, area in stats[2:]:
        checkbox = blackAndWhiteImage[y:y + h, x: x + w]
        sum_of_sums.append(np.sum(checkbox))
    if len(sum_of_sums) > 0:
        if len(checkbox_list) > np.argmax(sum_of_sums):
            vehicle_two_status = checkbox_list[np.argmax(sum_of_sums)]

    # Mode
    cv2.rectangle(pageTwo, (60, 1645), (690, 1690), (255, 0, 0), 2)
    mode_checkbox = pageTwo[1645:1690, 60:690]
    grayImage = cv2.cvtColor(mode_checkbox, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 127, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage

    stats = detect_checkbox(mode_checkbox)
    checkbox_list = ['Autonomous Mode', 'Conventional Mode']
    idx = 0
    sum_of_sums = []
    for x, y, w, h, area in stats[2:]:
        checkbox = blackAndWhiteImage[y:y + h, x: x + w]
        sum_of_sums.append(np.sum(checkbox))
    if len(sum_of_sums) > 0:
        if len(checkbox_list) > np.argmax(sum_of_sums):
            mode = checkbox_list[np.argmax(sum_of_sums)]
    #
    #     # PAGE THREE STARTED

    if pageThreeExist is not True:
        output = (
            s1_manu_name, s1_avt, s1_business_name, s1_telephone_number, s1_street_address,
            s1_city, s1_state, s1_zip, s2_date, s2_time, s2_vehicle_year, s2_make, s2_model,
            s2_license_plate_number, s2_vehicle_identification_number, s2_sviri, s2_address,
            s2_city, s2_country, s2_state, s2_zip, vehicle_one_status, s2_vehicles_involved,
            s2_driver_fullname, s2_driver_license_number, s2_driver_state, s2_driver_dob,
            s2_driver_insurance, s2_driver_policy_number, s2_naic_number, s2_policy_period,
            vehicle_one_damage, s3_vehicle_year, s3_model, s3_license_plate_number,
            s3_vehicle_identification_number, vehicle_two_status, s3_sviri, s3_vehicles_involved, s3_driver_fullname,
            s3_driver_license, s3_state, s3_driver_dob, s3_driver_insurance, s3_driver_policy_number,
            s3_naic_number, s3_policy_period, mode, ' '.join(description), vehicle_one_weather,
            vehicle_one_other_weather_choices, vehicle_two_weather, vehicle_two_other_weather_choices,
            vehicle_one_lighting, vehicle_one_other_lighting_choices, vehicle_two_lighting,
            vehicle_two_other_lighting_choices, vehicle_one_roadway_surface,
            vehicle_one_other_roadway_surface_choices, vehicle_two_roadway_surface,
            vehicle_two_other_roadway_surface_choices, vehicle_one_roadway_condition,
            vehicle_one_other_roadway_condition_choices, vehicle_two_roadway_condition,
            vehicle_two_other_roadway_condition_choices, vehicle_one_movement,
            vehicle_one_other_movement_choices, vehicle_two_movement,
            vehicle_two_other_movement_choices, vehicle_one_collision_type,
            vehicle_one_other_collision_type, vehicle_two_collision_type,
            vehicle_two_other_collision_type)

        ws.append(output)
        continue

    line_min_width = 20
    kernal_h = np.ones((1, line_min_width), np.uint8)

    # Weather

    # Vehicle 1
    vehicle_one_weather_img = pageThree[180:536, 456:557]
    grayImage = cv2.cvtColor(vehicle_one_weather_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    vehicle_one_weather_choices = []
    idx = 0
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_one_weather_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_one_weather_choices) > 0:
        vehicle_one_weather = vehicle_one_weather_choices[0]
        del vehicle_one_weather_choices[0]
        vehicle_one_other_weather_choices = ','.join(vehicle_one_weather_choices)

    # Vehicle 2
    cv2.rectangle(pageThree, (550, 179), (657, 537), (255, 0, 0), 1)
    vehicle_two_weather_img = pageThree[180:536, 551:656]
    grayImage = cv2.cvtColor(vehicle_two_weather_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    idx = 0
    vehicle_two_weather_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            cross_found = True
            if idx < len(checkbox_list):
                vehicle_two_weather_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1
    if len(vehicle_two_weather_choices) > 0:
        vehicle_two_weather = vehicle_two_weather_choices[0]
        del vehicle_two_weather_choices[0]
        vehicle_two_other_weather_choices = ','.join(vehicle_two_weather_choices)

    # Lighting

    # Vehicle 1
    cv2.rectangle(pageThree, (455, 580), (558, 870), (255, 0, 0), 1)
    vehicle_one_lighting_img = pageThree[581:869, 456:557]
    grayImage = cv2.cvtColor(vehicle_one_lighting_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    idx = 0
    vehicle_one_lighting_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_one_lighting_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_one_lighting_choices) > 0:
        vehicle_one_lighting = vehicle_one_lighting_choices[0]
        del vehicle_one_lighting_choices[0]
        vehicle_one_other_lighting_choices = ','.join(vehicle_one_lighting_choices)

    # Vehicle 2
    cv2.rectangle(pageThree, (550, 580), (657, 867), (255, 0, 0), 1)
    vehicle_two_lighting_img = pageThree[581:866, 551:656]
    grayImage = cv2.cvtColor(vehicle_two_lighting_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    idx = 0
    vehicle_two_lighting_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            cross_found = True
            if idx < len(checkbox_list):
                vehicle_two_lighting_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_two_lighting_choices) > 0:
        vehicle_two_lighting = vehicle_two_lighting_choices[0]
        del vehicle_two_lighting_choices[0]
        vehicle_two_other_lighting_choices = ','.join(vehicle_two_lighting_choices)

    # Roadway Surface
    # Vehicle 1
    cv2.rectangle(pageThree, (455, 910), (558, 1132), (255, 0, 0), 1)
    vehicle_one_roadway_surface_img = pageThree[911:1131, 456:557]
    grayImage = cv2.cvtColor(vehicle_one_roadway_surface_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D']
    idx = 0
    vehicle_one_roadway_surface_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_one_roadway_surface_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_one_roadway_surface_choices) > 0:
        vehicle_one_roadway_surface = vehicle_one_roadway_surface_choices[0]
        del vehicle_one_roadway_surface_choices[0]
        vehicle_one_other_roadway_surface_choices = ','.join(vehicle_one_roadway_surface_choices)

    # Vehicle 2
    cv2.rectangle(pageThree, (550, 910), (657, 1130), (255, 0, 0), 1)
    vehicle_two_roadway_surface_img = pageThree[911:1129, 551:656]
    grayImage = cv2.cvtColor(vehicle_two_roadway_surface_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    idx = 0
    vehicle_two_roadway_surface_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_two_roadway_surface_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_two_roadway_surface_choices) > 0:
        vehicle_two_roadway_surface = vehicle_two_roadway_surface_choices[0]
        del vehicle_two_roadway_surface_choices[0]
        vehicle_two_other_roadway_surface_choices = ','.join(vehicle_two_roadway_surface_choices)

    # Roadway Conditions
    # Vehicle 1
    cv2.rectangle(pageThree, (455, 1180), (558, 1653), (255, 0, 0), 1)
    vehicle_one_roadway_condition_img = pageThree[1181:1652, 456:557]
    grayImage = cv2.cvtColor(vehicle_one_roadway_condition_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    idx = 0
    vehicle_one_roadway_condition_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_one_roadway_condition_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_one_roadway_condition_choices) > 0:
        vehicle_one_roadway_condition = vehicle_one_roadway_condition_choices[0]
        del vehicle_one_roadway_condition_choices[0]
        vehicle_one_other_roadway_condition_choices = ','.join(vehicle_one_roadway_condition_choices)

    # Vehicle 2
    cv2.rectangle(pageThree, (550, 1183), (657, 1650), (255, 0, 0), 1)
    vehicle_two_roadway_condition_img = pageThree[1183:1649, 551:656]
    grayImage = cv2.cvtColor(vehicle_two_roadway_condition_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    idx = 0
    vehicle_two_roadway_condition_choices = []
    first = True
    for x, y, w, h, area in stats[2:]:
        if first:
            first = False
            continue
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        # cv2.imshow('', small_box)
        # cv2.waitKey(0)
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_two_roadway_condition_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_two_roadway_condition_choices) > 0:
        vehicle_two_roadway_condition = vehicle_two_roadway_condition_choices[0]
        del vehicle_two_roadway_condition_choices[0]
        vehicle_two_other_roadway_condition_choices = ','.join(vehicle_two_roadway_condition_choices)

    # Vehicle One Movement
    cv2.rectangle(pageThree, (1040, 178), (1140, 1131), (255, 0, 0), 1)
    vehicle_one_movement_img = pageThree[179:1130, 1041:1139]
    grayImage = cv2.cvtColor(vehicle_one_movement_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    vehicle_one_movement_choices = []
    idx = 0
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_one_movement_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_one_movement_choices) > 0:
        vehicle_one_movement = vehicle_one_movement_choices[0]
        del vehicle_one_movement_choices[0]
        vehicle_one_other_movement_choices = ','.join(vehicle_one_movement_choices)

    # Vehicle Two Movement
    cv2.rectangle(pageThree, (1132, 178), (1231, 1131), (255, 0, 0), 1)
    vehicle_two_movement_img = pageThree[179:1130, 1133:1230]
    grayImage = cv2.cvtColor(vehicle_two_movement_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
    idx = 0
    vehicle_two_movement_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_two_movement_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1

    if len(vehicle_two_movement_choices) > 0:
        vehicle_two_movement = vehicle_two_movement_choices[0]
        del vehicle_two_movement_choices[0]
        vehicle_two_other_movement_choices = ','.join(vehicle_two_movement_choices)

    # Collision Type

    # Vehicle 1
    cv2.rectangle(pageThree, (1040, 1180), (1140, 1654), (255, 0, 0), 1)
    vehicle_one_collision_type_img = pageThree[1181:1653, 1041:1139]
    grayImage = cv2.cvtColor(vehicle_one_collision_type_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    checkbox_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    idx = 0
    vehicle_one_collision_type_choices = []
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_one_collision_type_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1
    if len(vehicle_one_collision_type_choices) > 0:
        vehicle_one_collision_type = vehicle_one_collision_type_choices[0]
        del vehicle_one_collision_type_choices[0]
        vehicle_one_other_collision_type = ','.join(vehicle_one_collision_type_choices)

    # Vehicle 2
    cv2.rectangle(pageThree, (1133, 1180), (1230, 1654), (255, 0, 0), 1)
    vehicle_two_collision_type_img = pageThree[1181:1655, 1134:1231]
    grayImage = cv2.cvtColor(vehicle_two_collision_type_img, cv2.COLOR_BGR2GRAY)
    (thresh, blackAndWhiteImage) = cv2.threshold(grayImage, 230, 255, cv2.THRESH_BINARY)
    blackAndWhiteImage = 255 - blackAndWhiteImage
    img_bin_h = cv2.morphologyEx(blackAndWhiteImage, cv2.MORPH_OPEN, kernal_h)
    ret, labels, stats, centroids = cv2.connectedComponentsWithStats(img_bin_h, connectivity=8, ltype=cv2.CV_32S)
    if len(stats) > 1 and len(stats[1]) > 1:
        last_y = stats[1][1]
    vehicle_two_collision_type_choices = []
    idx = 0
    for x, y, w, h, area in stats[2:]:
        small_box = blackAndWhiteImage[last_y + 10:y - 10, x + 10: x + w - 10]
        if cv2.countNonZero(small_box) > 0:
            if idx < len(checkbox_list):
                vehicle_two_collision_type_choices.append(checkbox_list[idx])
        last_y = y
        idx += 1
    if len(vehicle_two_collision_type_choices) > 0:
        vehicle_two_collision_type = vehicle_two_collision_type_choices[0]
        del vehicle_two_collision_type_choices[0]
        vehicle_two_other_collision_type = ','.join(vehicle_two_collision_type_choices)

    output = (
        s1_manu_name, s1_avt, s1_business_name, s1_telephone_number, s1_street_address,
        s1_city, s1_state, s1_zip, s2_date, s2_time, s2_vehicle_year, s2_make, s2_model,
        s2_license_plate_number, s2_vehicle_identification_number, s2_sviri, s2_address,
        s2_city, s2_country, s2_state, s2_zip, vehicle_one_status, s2_vehicles_involved,
        s2_driver_fullname, s2_driver_license_number, s2_driver_state, s2_driver_dob,
        s2_driver_insurance, s2_driver_policy_number, s2_naic_number, s2_policy_period,
        vehicle_one_damage, s3_vehicle_year, s3_model, s3_license_plate_number,
        s3_vehicle_identification_number, vehicle_two_status, s3_sviri, s3_vehicles_involved, s3_driver_fullname,
        s3_driver_license, s3_state, s3_driver_dob, s3_driver_insurance, s3_driver_policy_number,
        s3_naic_number, s3_policy_period, mode, ' '.join(description), vehicle_one_weather,
        vehicle_one_other_weather_choices, vehicle_two_weather, vehicle_two_other_weather_choices,
        vehicle_one_lighting, vehicle_one_other_lighting_choices, vehicle_two_lighting,
        vehicle_two_other_lighting_choices, vehicle_one_roadway_surface,
        vehicle_one_other_roadway_surface_choices, vehicle_two_roadway_surface,
        vehicle_two_other_roadway_surface_choices, vehicle_one_roadway_condition,
        vehicle_one_other_roadway_condition_choices, vehicle_two_roadway_condition,
        vehicle_two_other_roadway_condition_choices, vehicle_one_movement,
        vehicle_one_other_movement_choices, vehicle_two_movement,
        vehicle_two_other_movement_choices, vehicle_one_collision_type,
        vehicle_one_other_collision_type, vehicle_two_collision_type,
        vehicle_two_other_collision_type)

    ws.append(output)
wb.save('Cal_DMV_AV_Dataset.xlsx')

now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)

# vehicle involved (DONE)
# PDF TO JPG   (DONE)
# Extract to excel (DONE)
# Multiple file picks (DONE)
