import phonenumbers
from phonenumbers import carrier, timezone, geocoder
from phonenumbers.phonenumberutil import number_type

number = "955063537"
print(carrier._is_mobile(number_type(phonenumbers.parse(number))))
print(phonenumbers.parse("7955063537"))

my_number = phonenumbers.parse("7955063537")
print(1)
print(carrier.name_for_number(my_number, "en"))
print(2)
print(timezone.time_zones_for_number(my_number))
print(3)
print(geocoder.description_for_number(my_number, "en"))
print(4)
print(phonenumbers.is_valid_number(my_number))
print(5)
print(phonenumbers.is_possible_number(my_number))