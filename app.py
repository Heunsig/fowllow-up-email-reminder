from __future__ import print_function
import time
import sys
# import threading
# import json
import schedule
import pydash
import Smtp
from GoogleSpreadsheet import GoogleSpreadsheet
from datetime_helper import Easydate


# PROGRAM_FINISHED = True


def analzeNewProspective(values, indexs):
  # global YESTERDAY

  students = []
  sources = []
  campuses = []
  nationalities = []

  for value in values:
    try:
      # if date.strptime(value[indexs["respondDate"]], '%m/%d/%Y').strftime('%m/%d/%y') == Datetime.getToday():
      if Easydate.convertDate(value[indexs["respondDate"]]) == Easydate.today():
        students.append([value[indexs["name"]]])
        sources.append(value[indexs["source"]])
        campuses.append(value[indexs["campus"]])
        nationalities.append(value[indexs["nationality"]])
    except (ValueError, IndexError):
      pass

  return {
      "students": len(students),
      "sources": pydash.count_by(sources, lambda x: x.lower()),
      "campuses": pydash.count_by(campuses, lambda x: x.replace(' ', '').lower()),
      "nationalities": pydash.count_by(nationalities, lambda x: x.lower()),
  }


def reapeatGettingValues(GoogleAPI):
  values = GoogleAPI.loadValues()

  if not values:
    print('No data found.')
  else:
    print('You\'ve successfully gotten spreadsheet data.')
    contactIndex = values[0].index('Contact')
    campusIndex = values[0].index('Campus')
    nameIndex = values[0].index('Name')
    SourceIndex = values[0].index('Source')
    phoneIndex = values[0].index('PHONE NUMBER')
    emailIndex = values[0].index('Email')
    nationalityIndex = values[0].index('Nationality')
    detailsIndex = values[0].index('Details')
    respondDateIndex = values[0].index('Date Of respond')

    firstDateIndex = values[0].index('1st follow up')
    secondDateIndex = values[0].index('2nd follow up')
    thirdDateIndex = values[0].index('3rd follow up')

    indexs = {
        "contact": contactIndex,
        "campus": campusIndex,
        "name": nameIndex,
        "source": SourceIndex,
        "phone": phoneIndex,
        "email": emailIndex,
        "nationality": nationalityIndex,
        "details": detailsIndex,
        "respondDate": respondDateIndex
    }

    firstFolloups = appendFollowupList(values, firstDateIndex, indexs)
    secondFolloups = appendFollowupList(values, secondDateIndex, indexs)
    thirdFolloups = appendFollowupList(values, thirdDateIndex, indexs)

    analysisData = analzeNewProspective(values, indexs)

    smtp = Smtp.Smtp('heun3344@gmail.com', 'D8t44m5b@#', 'smtp.gmail.com', 587)
    smtp.login()

    sendAnalysisDataEmail(smtp, analysisData)

    sendFollowUpEmail(smtp, firstFolloups, 'The First Follow up')
    sendFollowUpEmail(smtp, secondFolloups, 'The Second Follow up')
    sendFollowUpEmail(smtp, thirdFolloups, 'The Third Follow up')

    smtp.quit()

    # print(analzeNewProspective(values, indexs))
    # f = open('guru.json', 'w')
    # f.write(json.dumps(firstFolloups, separators=(',', ':')))
    # f.close()


def appendFollowupList(arr, dateIndex, indexs):
  # global TODAY
  dates = []
  # today = datetime.now().strftime('%m/%d/%y')
  # today = '04/18/19'

  for item in arr:
    try:
      if Easydate.convertDate(item[dateIndex]) == Easydate.today():
        dates.append({
            "contact": item[indexs["contact"]],
            "campus": item[indexs["campus"]],
            "name": item[indexs["name"]],
            "source": item[indexs["source"]],
            "phone": item[indexs["phone"]],
            "email": item[indexs["email"]],
            "nationality": item[indexs["nationality"]],
            "details": item[indexs["details"]],
            "date": item[dateIndex]
        })
    except IndexError:
      pass
      # try:
      #   try:
      #     if Easydate.convertDate(item[dateIndex], '%m/%d/%y').strftime('%m/%d/%y') == Easydate.today():

      #       dates.append({
      #           "contact": item[indexs["contact"]],
      #           "campus": item[indexs["campus"]],
      #           "name": item[indexs["name"]],
      #           "source": item[indexs["source"]],
      #           "phone": item[indexs["phone"]],
      #           "email": item[indexs["email"]],
      #           "nationality": item[indexs["nationality"]],
      #           "details": item[indexs["details"]],
      #           "date": item[dateIndex]
      #       })
      #   except ValueError:
      #     try:
      #       if date.strptime(item[dateIndex], '%m/%d/%Y').strftime('%m/%d/%y') == TODAY:

      #         dates.append({
      #             "contact": item[indexs["contact"]],
      #             "campus": item[indexs["campus"]],
      #             "name": item[indexs["name"]],
      #             "source": item[indexs["source"]],
      #             "phone": item[indexs["phone"]],
      #             "email": item[indexs["email"]],
      #             "nationality": item[indexs["nationality"]],
      #             "details": item[indexs["details"]],
      #             "date": date.strptime(item[dateIndex], '%m/%d/%Y').strftime('%m/%d/%y')
      #         })
      #     except ValueError:
      #       pass
      # except (IndexError):
      #   pass

  return dates


def sendFollowUpEmail(email_service, follows, subject):
  for follow in follows[0:1]:
    text = '''
    Hello %s,

    Today (%s) you need to do %s.
    Name: %s
    Nationality: %s
    Campus: %s
    Phone: %s
    Email: %s
    Details: %s

    Please don't forget sending a follow-up email for our students.
    Thank you.
    ''' % (follow["contact"], follow["date"], subject, follow["name"], follow["nationality"], follow["campus"], follow["phone"], follow["email"], follow["details"])
    print(text)

    email_service.send(to='heun3344@gmail.com', subject=subject, content=text)

    time.sleep(2)


def sendAnalysisDataEmail(email_service, data):
  sources = ''
  for key, count in data["sources"].items():
    sources += key + " : " + str(count) + ' '

  campuses = ''
  for key, count in data["campuses"].items():
    campuses += key + " : " + str(count) + ' '

  nationalities = ''
  for key, count in data["nationalities"].items():
    nationalities += key + " : " + str(count) + ' '

  text = '''
    Hello everyone,

    Yesterday we had some prospective students.
    Total: %s
    Details------
    - Sources: %s,
    - campuses: %s,
    - nationailties: %s
  ''' % (data["students"], sources, campuses, nationalities)

  email_service.send(to='heun3344@gmail.com', subject='Yesterday summary', content=text)


def main():
  global PROGRAM_FINISHED

  try:
    print('Program Started!!!')

    GoogleAPI = GoogleSpreadsheet(
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'],
        spreadsheet_id='1thT4M0AFcBucct1PLk7vL29Fhrxz1YnxOAq-06d64V8',
        range_name='April 19')

    print('Authorizing...')
    if GoogleAPI.authorize():
      print('Authorized!!!')

    schedule.every(5).seconds.do(reapeatGettingValues, GoogleAPI=GoogleAPI)
    # schedule.every().minute.at(":02").do(reapeatGettingValues, GoogleAPI=GoogleAPI)

    while True:
      schedule.run_pending()
      time.sleep(1)

  except KeyboardInterrupt:
    sys.exit()

  print('Program finished!!!')
  sys.exit()


if __name__ == '__main__':
  # # test = Good()
  # # print(Good.yesterday())
  # # print(Good.today())
  # # print(Datetime.getToday())
  # print(Good.convertDate('04/21/19') == Good.today())
  # # print(datetime.strptime('04/04/19', '%m/%d/%y').date())
  main()
