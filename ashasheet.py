import sys
import csv
from datetime import datetime
from easygui import *

# Splash
msgbox("Begin by selecting a CSV")

# Load input file
filename = fileopenbox(msg="Select a CSV", title="Source File", default='*.csv', filetypes=['*.csv'])

# Begin processing file
with open(filename, 'r') as infile:
    reader = csv.DictReader(infile)

    headers = reader.fieldnames

    # Display fields for confirmation
    id_key = choicebox(msg='Which field contains ASHA IDs', title='ASHA ID', choices=headers)
    ceu_key = choicebox(msg='Which field contains CEU requests', title='CEU Request', choices=headers)

    attendee_count = 0
    ceu_users = []
    next(reader, None)
    for row in reader:
        attendee_count = attendee_count + 1
        if row[ceu_key].lower() == 'yes':
            ceu_users.append(row)

    # Iterate through the 'Date Attending' column until we find one that can be used
    event_time = None
    infile.seek(0)
    while True:
        try:
            event_time = datetime.strptime(next(reader)['Date Attending'], '%b %d, %Y at %I:%M %p')
            break
        except ValueError:
            pass

    # Enter Course Offering data
    msg = "Enter your personal information"
    title = "Course Offering"
    field_names = ["Provider_code", "CourseNumber", "Part_forms",
                   "Offering_complete_date", "number_attending", "Partial_credit"]
    field_values = ["ABCD", "1234567", len(ceu_users), event_time.strftime('%m%d%y'), attendee_count, "N"]
    field_values = multenterbox("Enter Course Offering Data", "Course offering", field_names, field_values)

    # Ensure proper padding
    field_values[1] = field_values[1].zfill(7)
    field_values[2] = field_values[2].zfill(4)
    field_values[4] = field_values[4].zfill(5)

    # Create output files
    out_fields = ["Form_type"] + field_names
    ref_fields = ["Customer_Name", "ASHA_ID", "Address1", "Address2", "Address3", "City", "State", "ZIP", "Country",
                  "Primary_Phone", "Email", "Provider_RefNum"]

    base_filename = field_values[0] + datetime.now().strftime('%Y%m%d') + "01"

    # Produce primary output file
    out_filename = filesavebox(msg="Select filename to save", title="Output file", default=base_filename + ".csv",
                               filetypes=['*.csv'])
    with open(out_filename, 'a') as outfile:
        out_writer = csv.DictWriter(outfile, fieldnames=out_fields)
        out_writer.writeheader()
        out_writer.writerow(dict(zip(out_fields, ["A1"] + field_values)))

        # Copy over data from input file
        infile.seek(0)
        for row in ceu_users:
            # writes the reordered rows to the new file
            in_row = dict(zip(out_fields, [
                "P1",
                row['Last Name'],
                '',
                str(row[id_key]).zfill(8),
                ''
            ]))
            out_writer.writerow(in_row)

    # Produce user reference file
    ref_filename = filesavebox(msg="Select filename to save", title="Reference file", default=base_filename + "ref.csv",
                               filetypes=['*.csv'])
    with open(ref_filename, 'a') as reffile:
        ref_writer = csv.DictWriter(reffile, fieldnames=ref_fields)
        ref_writer.writeheader()

        # Copy over data from input file
        for row in ceu_users:
            # writes the reordered rows to the new file
            in_row = dict(zip(ref_fields, [
                row['Last Name'] + ' ' + row['First Name'],
                row[id_key],
                row["Home Address 1"],
                row["Home Address 2"],
                '',
                row["Home City"],
                row["Home State"],
                row["Home Zip"],
                row["Home Country"],
                row["Cell Phone"],
                row["Email"],
                row["Attendee #"]
            ]))
            ref_writer.writerow(in_row)







'''
with open(filename, 'r') as infile, open('asha_Formatted.csv', 'a') as outfile:
    # output dict needs a list for new column ordering
    fieldnames = ['A', 'C', 'D', 'E', 'B']
    writer = csv.DictWriter(outfile, fieldnames=fieldnames)
    # reorder the header first
    writer.writeheader()
    for row in csv.DictReader(infile):
        # writes the reordered rows to the new file
        writer.writerow(row)


while 1:
    msgbox("Hello, world!")

    msg ="What is your favorite flavor?"
    title = "Ice Cream Survey"
    choices = ["Vanilla", "Chocolate", "Strawberry", "Rocky Road"]
    choice = choicebox(msg, title, choices)

    # note that we convert choice to string, in case
    # the user cancelled the choice, and we got None.
    msgbox("You chose: " + str(choice), "Survey Result")

    msg = "Do you want to continue?"
    title = "Please Confirm"
    if ccbox(msg, title):     # show a Continue/Cancel dialog
        pass  # user chose Continue
    else:
        sys.exit(0)           # user chose Cancel
'''