import xlsxwriter
from openpyxl import Workbook
import requests
import re
import sys
from pprint import pprint

# Global variable that stores the API token
at = "6vWZ9c9kuwSdZHlrnDDZjFnIMjIVFRG3PltznQEA"
workbook = xlsxwriter.Workbook('Dump.xlsx')
worksheet = workbook.add_worksheet()

####### Interact with user for necessary parameters for analysis

# Initial menu for fetching the API token and the desired group to be analyzed
def menu():
    global at
    print('If you have not done so already, go to the following website to receive your API token: ' +
          'https://dev.groupme.com/. When signing up, it does not matter what you put for the callback URL')
    print("Here are your ten most recent groups:")
    groups_data = print_all_groups_with_number_beside_each()
    try:
        group_number = int(input("Enter the number of the group you would like to analyze:"))
        group_id = get_group_id(groups_data, group_number)
        prepare_analysis_of_group(groups_data, group_id)
    except ValueError:
        print("Not a number")
        worksheet.write(0, 4, 'ValueError')
        workbook.close()



# List all groups open to user
def print_all_groups_with_number_beside_each():
    response = requests.get('https://api.groupme.com/v3/groups?token=' + at)
    data = response.json()
    if len(data['response']) == 0:
        print("You are not part of any groups.")
        return
    for i in range(len(data['response'])):
        group = data['response'][i]['name']
        print(str(i) + "\'" + group + "\'")
    return data


####### Methods for getting group information

# Find the name of groups
def get_group_name(groups_data, group_id):
    i = 0
    while True:
        if group_id == groups_data['response'][i]['group_id']:
            return groups_data['response'][i]['name']
        i += 1


# Find the ID of the selected group to proceed with analysis
def get_group_id(groups_data, group_number):
    group_id = groups_data['response'][group_number]['id']
    return group_id


def get_number_of_messages_in_group(groups_data, group_id):
    i = 0
    while True:
        if group_id == groups_data['response'][i]['group_id']:
            return groups_data['response'][i]['messages']['count']
        i += 1


def get_group_members(groups_data, group_id):
    i = 0
    while True:
        if group_id == groups_data['response'][i]['group_id']:
            return groups_data['response'][i]['members']
        i += 1


####### Analyzing group messages

# Fetching basic metrics of the group
def prepare_analysis_of_group(groups_data, group_id):
    # Return basic information of the group
    group_name = get_group_name(groups_data, group_id)
    number_of_messages = get_number_of_messages_in_group(groups_data, group_id)
    print("Analyzing " + str(number_of_messages) + " messages from " + group_name)
    # Map the users
    members_of_group_data = get_group_members(groups_data, group_id)
    user_dictionary = prepare_user_dictionary(members_of_group_data)
    # Analyze the group's messages
    user_id_mapped_to_user_data = analyze_group(group_id, user_dictionary, number_of_messages)
    # Return the data
    display_data(user_id_mapped_to_user_data)


# Map users
def prepare_user_dictionary(members_of_group_data):
    user_dictionary = {}
    i = 0
    while True:
        try:
            # Get information of the user
            user_id = members_of_group_data[i]['user_id']
            nickname = members_of_group_data[i]['nickname']
            user_dictionary[user_id] = [nickname, 0.0, 0.0, 0.0]
            # Optional metrics that can be measured for each user:
            # [0] = nickname, 
            # [1] = total messages sent in group, like count, 
            # [2] = likes per message,
            # [3] = average likes received per message, 
            # [4] = total words sent, 
            # [5] = dictionary of likes received from each member
            # [6] = dictionary of shared likes, 
            # [7] = total likes given

        except IndexError:
            return user_dictionary
        i += 1
    return user_dictionary


# Analyzing the messages
def analyze_group(group_id, user_id_mapped_to_user_data, number_of_messages):
    response = requests.get('https://api.groupme.com/v3/groups/' + group_id + '/messages?token=' + at)
    data = response.json()
    message_with_only_alphanumeric_characters = ''
    message_id = 0
    iterations = 0.0
    while True:
        for i in range(20):  # in range of 20 because API sends 20 messages at once
            try:
                iterations += 1
                name = data['response']['messages'][i]['name']  # grabs name of sender
                message = data['response']['messages'][i]['text']  # grabs text of message
                # print(message)
                try:
                    #  strips out special characters
                    message_with_only_alphanumeric_characters = re.sub(r'\W+', ' ', str(message))
                except ValueError:
                    pass  # this is here to catch errors when there are special characters in the message e.g. emoticons
                sender_id = data['response']['messages'][i]['sender_id']  # grabs sender id
                list_of_favs = data['response']['messages'][i]['favorited_by']  # grabs list of who favorited message
                length_of_favs = len(list_of_favs)  # grabs number of users who liked message

                # grabs the number of words in message
                number_of_words_in_message = len(re.findall(r'\w+', str(message_with_only_alphanumeric_characters)))

                if sender_id not in user_id_mapped_to_user_data.keys():
                    user_id_mapped_to_user_data[sender_id] = [name, 0.0, 0.0, 0.0, 0.0, {}, {}, 0.0]

                # this if statement is here to fill the name in for the case where a user id liked a message but had
                # yet been added to the dictionary
                if user_id_mapped_to_user_data[sender_id][0] == '':
                    user_id_mapped_to_user_data[sender_id][0] = name



                for user_id in list_of_favs:
                    for user_id_inner in list_of_favs:
                        if user_id not in user_id_mapped_to_user_data.keys():
                            # leave name blank because this means a user is has liked a message but has yet to be added
                            # to the dictionary. So leave the name blank until they send their first message.
                            user_id_mapped_to_user_data[user_id] = ['', 0.0, 0.0, 0.0, 0.0, {}, {}, 0.0]

                user_id_mapped_to_user_data[sender_id][2] += length_of_favs

            except IndexError:
                print("COMPLETE")
                print
                for key in user_id_mapped_to_user_data:
                    try:
                        user_id_mapped_to_user_data[key][3] = user_id_mapped_to_user_data[key][2] / \
                                                              user_id_mapped_to_user_data[key][1]
                    except ZeroDivisionError:  # for the case where the user has sent 0 messages
                        user_id_mapped_to_user_data[key][3] = 0
                return user_id_mapped_to_user_data

        if i == 19:
            message_id = data['response']['messages'][i]['id']
            remaining = iterations / number_of_messages
            remaining *= 100
            remaining = round(remaining, 2)
            print(str(remaining) + ' percent done')

        payload = {'before_id': message_id}
        response = requests.get('https://api.groupme.com/v3/groups/' + group_id + '/messages?token=' + at,
                                params=payload)
        data = response.json()


####### Information rendering/parsing

def display_data(user_id_mapped_to_user_data):
    counter = 0
    col = 0
    for key in user_id_mapped_to_user_data:
        counter += 1
        try:
            worksheet.write(counter, col, user_id_mapped_to_user_data[key][0])
            worksheet.write(counter, col + 1, user_id_mapped_to_user_data[key][2])
        except KeyError:
            print("Somthing Happened :/" + key)
            worksheet.write(0, 4, 'KeyError')
            workbook.close()


# Uncomment this line below to view the raw dictionary


# Initiate program
menu()
worksheet.write(0, 4, 'No Errors!')
workbook.close()

