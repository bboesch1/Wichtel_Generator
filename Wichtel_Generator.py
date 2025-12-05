"""
This is a Wichtel Generator built for Christmas 2024
Date: 13.12.2024
Author: Benno Bösch
Version: 0.0.1

Update: Added email functionality
Version 0.0.2
Date 05.12.2024
"""

import random
import win32com.client as win32


def get_participants():
    participants_email = {}
    participants = []

    nbr_participants = int(input("Enter the number of participants: "))
    for i in range(nbr_participants):
        participant = input("Enter participant " + str(i+1) + " name: ")
        participant_email = input('Enter participants: ' + participant + ' email address: ')

        participants_email[participant] = participant_email
        participants.append(participant)

    return participants, participants_email

def check_draw_validity(participants, allocation_list):
    valid = True
    for i in range(len(participants)):
        if participants[i] == allocation_list[i]:
            print(participants[i], " draws ", allocation_list[i])
            valid = False
            print(valid)
    return valid


def wichtel_allocation(participants):
    allocation_list = []
    drawing_pot = participants.copy()
    for i in range(len(participants)):
        name = random.choice(drawing_pot)
        if name == participants[i]:
            name = random.choice(drawing_pot)
        drawing_pot.remove(name)
        allocation_list.append(name)
    return allocation_list

def save_wichtel_allocation(participants, allocation_list):
    for i in range(len(participants)):
        f = open(participants[i]+".txt", "w")
        f.write("Du wichtlisch a de folgende Person:\n")
        f.write(allocation_list[i]+"\n")
        f.write("Ganz viel Erfolg und wiiterhin schöni Adventsziit :)\n")
    return

def send_emails(participants, participants_email, allocation_list):
    outlook = win32.Dispatch('outlook.application')
    for participant in participants:
        mail = outlook.CreateItem(0)
        mail.To = participants_email[participant]
        mail.Subject = 'Wichteln 2025 Zuteilung TEST'
        mail.Body = f"""
                    Hoi {participant} \n
                    Du döfsch uf die Wiehnachte a {allocation_list[0]} wichtle. Ich wünsche dier viel Spass uf de Suechi, may the force be with you.

                    Nomal en churzi zämmefassig vo de Spielregle:
                        Maximalbetrag isch 50 CHF
                        Möglichst unerkenntlich verpacke, damit mier au no chli chönd rate am D-Day

                    Jetzt wünsch ich dier no schöni Wiehnachtsziit und bis gli
                    Din Wichtelgenerator
                    """
        mail.Send()



def main():
    print("This is a Wichtel generator. First please enter all participants")
    participants, participants_email = get_participants()

    validity = False

    while not validity:
        allocation_list = wichtel_allocation(participants)
        validity = check_draw_validity(participants, allocation_list)

    save_wichtel_allocation(participants, allocation_list)

    send_emails(participants, participants_email, allocation_list)
    #print(participants)
    #print(allocation_list)
    return

if __name__ == "__main__":
    main()
