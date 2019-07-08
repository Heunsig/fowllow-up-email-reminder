from __future__ import print_function
import time
import sys
import schedule
import pydash
import Smtp
from GoogleSpreadsheet import GoogleSpreadsheet
from Easydate import Easydate
import datetime


def analzeNewProspective(values, indexs):
  students = []
  sources = []
  campuses = []
  nationalities = []

  for value in values:
    try:
      # if Easydate.convertDate(value[indexs["respondDate"]]) == Easydate.today():
      if Easydate.convertDate(value[indexs["respondDate"]]) == Easydate.yesterday():
        # print(type(Easydate.yesterday()))
        # if Easydate.convertDate(value[indexs["respondDate"]]) == datetime.date(2019, 4, 22):
        if value[indexs["name"]]:
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

    print('Sending emails at %s' % Easydate.now())
    smtp = Smtp.Smtp('heun3344@gmail.com', 'D8t44m5b@#', 'smtp.gmail.com', 587)
    smtp.login()

    sendAnalysisDataEmail(smtp, analysisData)
    print('The analysis data email finished')

    sendFollowUpEmail(smtp, firstFolloups, 'The First Follow up')
    print('The first follow up email finished')
    sendFollowUpEmail(smtp, secondFolloups, 'The Second Follow up')
    print('The second follow up email finished')
    sendFollowUpEmail(smtp, thirdFolloups, 'The Third Follow up')
    print('The third follow up email finished')

    smtp.quit()

    print('Sending emails finished!!!')
    print('Waiting for the next transcation!!!')


def appendFollowupList(arr, dateIndex, indexs):
  dates = []

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

  return dates


def sendFollowUpEmail(email_service, follows, subject):
  if follows:
    for follow in follows:
      # print('tes', follow["contact"].lower())
      if follow["contact"].lower() == 'alex':
        print('a')
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

        email_service.send(to='alex@mliesl.edu', subject=subject, content=text)

        time.sleep(2)


def sendAnalysisDataEmail(email_service, data):
  if data:
    sources = ''
    for key, count in data["sources"].items():
      sources += key + ": " + str(count) + ' '

    campuses = ''
    for key, count in data["campuses"].items():
      campuses += key + ": " + str(count) + ' '

    nationalities = ''
    for key, count in data["nationalities"].items():
      nationalities += key + ": " + str(count) + ' '

    text = '''
      Hello everyone,

      Yesterday we had some prospective students.
      Total: %s
      Details------
      - Sources: %s,
      - campuses: %s,
      - nationailties: %s
    ''' % (data["students"], sources, campuses, nationalities)
  else:
    text = '''
      No prospective students.
    '''

  email_service.send(to='heun3344@gmail.com', subject='Yesterday summary', content=text)


def main():
  try:
    print('Program Started!!!')

    GoogleAPI = GoogleSpreadsheet(
        scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'],
        spreadsheet_id='1thT4M0AFcBucct1PLk7vL29Fhrxz1YnxOAq-06d64V8',
        range_name='April 19')

    print('Authorizing...')
    if GoogleAPI.authorize():
      print('Authorized!!!')

    # schedule.every(30).seconds.do(reapeatGettingValues, GoogleAPI=GoogleAPI)
    schedule.every(2).seconds.do(reapeatGettingValues, GoogleAPI=GoogleAPI)
    # schedule.every().hour.at(":52").do(reapeatGettingValues, GoogleAPI=GoogleAPI)

    print('Waiting for the next transcation!!!')
    while True:
      schedule.run_pending()
      time.sleep(10)

  except KeyboardInterrupt:
    sys.exit()

  print('Program finished!!!')
  sys.exit()


if __name__ == '__main__':
  main()
