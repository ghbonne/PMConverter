from objects.agenda import Agenda
from datetime import datetime, timedelta

agenda = Agenda()


begindates = [None]*5
enddates = [None]*5
begindates[0] = datetime(day=9, month=2, year=2007, hour=8)
enddates[0] = agenda.get_end_date(begindates[0], 10)
print(enddates[0])

begindates[1]= datetime(day=9, month=2, year=2007, hour=8)
enddates[1] = agenda.get_end_date(begindates[1], 12)
print(enddates[1])

begindates[2]= datetime(day=28, month=2, year=2007, hour=8)
enddates[2] = agenda.get_end_date(begindates[2], 13)
print(enddates[2])

begindates[3]= datetime(day=28, month=2, year=2007, hour=8)
enddates[3] = agenda.get_end_date(begindates[3], 12)
print(enddates[3])

begindates[4]= datetime(day=16, month=3, year=2007, hour=8)
enddates[4] = agenda.get_end_date(begindates[4], 1)
print(enddates[4])


print("min")
print(min(begindates))
print("max")
print(max(enddates))

print(agenda.get_time_between(min(begindates), max(enddates)))