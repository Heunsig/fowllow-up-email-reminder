from datetime import datetime, date, timedelta


class Easydate:
  @staticmethod
  def today():
    return date.today()

  @staticmethod
  def yesterday():
    return (date.today() - timedelta(1))

  @staticmethod
  def now():
    return datetime.now()

  @staticmethod
  def convertDate(date):
    try:
      return datetime.strptime(date, '%m/%d/%y').date()

    except ValueError:
      try:
        return datetime.strptime(date, '%m/%d/%Y').date()

      except ValueError:
        try:
          return datetime.strptime(date, '%m/%d').date().replace(year=Easydate.today().year)
        except ValueError:
          pass
