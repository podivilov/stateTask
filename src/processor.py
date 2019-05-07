#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import uuid
import pytz
import xlrd
from xml.dom import minidom
from datetime import datetime

# Устанавливаем стандартную кодировку
reload(sys)
sys.setdefaultencoding('utf8')

# Полезные переменные
local_tz = pytz.timezone('Europe/Moscow')

# Полезные функции
def utc_to_local(utc_dt):
    local_dt = utc_dt.replace(tzinfo=pytz.utc).astimezone(local_tz)
    return local_tz.normalize(local_dt)

def insert_str(string, str_to_insert, index):
    return string[:index] + str_to_insert + string[index:]

# Открываем рабочий файл
workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_index(0)

# Создаём коренной элемент
doc = minidom.Document()
root = doc.createElement('ns2:stateTask640r')
doc.appendChild(root)

# Header
header = doc.createElement('header')
root.appendChild(header)

# ID
id = doc.createElement('id')
id.appendChild(doc.createTextNode(str(uuid.uuid4())))
header.appendChild(id)

# createDateTime
createDateTime = doc.createElement('createDateTime')
createDateTime.appendChild(doc.createTextNode(str(utc_to_local(datetime.utcnow()).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + "+03:00")))
header.appendChild(createDateTime)

# ns2:body
ns2_body = doc.createElement('ns2:body')
root.appendChild(ns2_body)

# ns2:position
ns2_position = doc.createElement('ns2:position')
ns2_body.appendChild(ns2_position)

# positionId
positionId = doc.createElement('positionId')
positionId.appendChild(doc.createTextNode(str(uuid.uuid4())))
ns2_position.appendChild(positionId)

# changeDate
changeDate = doc.createElement('changeDate')
changeDate.appendChild(doc.createTextNode(str(utc_to_local(datetime.utcnow()).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + "+03:00")))
ns2_position.appendChild(changeDate)

# placer
placer = doc.createElement('placer')
ns2_position.appendChild(placer)

# regNum
placer_regNum = doc.createElement('regNum')
placer_regNum.appendChild(doc.createTextNode('462D1140'))
placer.appendChild(placer_regNum)

# fullName
placer_fullName = doc.createElement('fullName')
placer_fullName.appendChild(doc.createTextNode('ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ МОСКОВСКОЙ ОБЛАСТИ «КОЛЛЕДЖ «КОЛОМНА»'))
placer.appendChild(placer_fullName)

# inn
placer_inn = doc.createElement('inn')
placer_inn.appendChild(doc.createTextNode('5022049898'))
placer.appendChild(placer_inn)

# kpp
placer_kpp = doc.createElement('kpp')
placer_kpp.appendChild(doc.createTextNode('502201001'))
placer.appendChild(placer_kpp)

# initiator
initiator = doc.createElement('initiator')
ns2_position.appendChild(initiator)

# regNum
initiator_regNum = doc.createElement('regNum')
initiator_regNum.appendChild(doc.createTextNode('462D1140'))
initiator.appendChild(initiator_regNum)

# fullName
initiator_fullName = doc.createElement('fullName')
initiator_fullName.appendChild(doc.createTextNode('ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ МОСКОВСКОЙ ОБЛАСТИ «КОЛЛЕДЖ «КОЛОМНА»'))
initiator.appendChild(initiator_fullName)

# inn
initiator_inn = doc.createElement('inn')
initiator_inn.appendChild(doc.createTextNode('5022049898'))
initiator.appendChild(initiator_inn)

# kpp
initiator_kpp = doc.createElement('kpp')
initiator_kpp.appendChild(doc.createTextNode('502201001'))
initiator.appendChild(initiator_kpp)

# versionNumber
versionNumber = doc.createElement('versionNumber')
versionNumber.appendChild(doc.createTextNode('0'))
ns2_position.appendChild(versionNumber)

# now
now = datetime.now()

# reportYear
reportYear = doc.createElement('reportYear')
reportYear.appendChild(doc.createTextNode(str(now.year)))
ns2_position.appendChild(reportYear)

# financialYear
financialYear = doc.createElement('financialYear')
financialYear.appendChild(doc.createTextNode(str(now.year)))
ns2_position.appendChild(financialYear)

# nextFinancialYear
nextFinancialYear = doc.createElement('nextFinancialYear')
nextFinancialYear.appendChild(doc.createTextNode(str(now.year + 1)))
ns2_position.appendChild(nextFinancialYear)

# planFirstYear
planFirstYear = doc.createElement('planFirstYear')
planFirstYear.appendChild(doc.createTextNode(str(now.year + 1)))
ns2_position.appendChild(planFirstYear)

# planLastYear
planLastYear = doc.createElement('planLastYear')
planLastYear.appendChild(doc.createTextNode(str(now.year + 2)))
ns2_position.appendChild(planLastYear)

### РАЗДЕЛ №1 ###

# service
service = doc.createElement('service')
ns2_position.appendChild(service)

# service_name
service_name = doc.createElement('name')
service_name.appendChild(doc.createTextNode('Реализация образовательных программ среднего профессионального образования - программ подготовки специалистов среднего звена'))
service.appendChild(service_name)

# type_name
type_name = doc.createElement('type')
type_name.appendChild(doc.createTextNode('S'))
service.appendChild(type_name)

# ordinalNumber
ordinalNumber = doc.createElement('ordinalNumber')
ordinalNumber.appendChild(doc.createTextNode('1'))
service.appendChild(ordinalNumber)

# category
category = doc.createElement('category')
service.appendChild(category)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('Физические лица, имеющие основное общее образование'))
category.appendChild(name)

# renderEnactment
renderEnactment = doc.createElement('renderEnactment')
service.appendChild(renderEnactment)

# type
type = doc.createElement('type')
type.appendChild(doc.createTextNode('Закон'))
renderEnactment.appendChild(type)

# author
author = doc.createElement('author')
renderEnactment.appendChild(author)

# fullName
fullName = doc.createElement('fullName')
fullName.appendChild(doc.createTextNode('Федеральный закон'))
author.appendChild(fullName)

# date
date = doc.createElement('date')
date.appendChild(doc.createTextNode('2012-12-29+04:00'))
renderEnactment.appendChild(date)

# number
number = doc.createElement('number')
number.appendChild(doc.createTextNode('273-фз'))
renderEnactment.appendChild(number)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('«Об образовании в Российской Федерации»'))
renderEnactment.appendChild(name)

#
# Алгоритм обработки строк таблицы
# Для извлечения значения из ячейки по индексу [stroka, stolbec] используется worksheet.cell(stroka, stolbec).value
#
# 1. Генерация сущностей типа <volumeIndex>
#    Сущность состоит из следующих сущностей: <index>, <deviation> (постоянное значение) и <valueYear>
#
#    Сущность <index> содержит в себе сущность <regNum>, которая отвечает за регистрационный номер
#    и извлекается из нулевой ячейки [0, stolbec]. За окончанием списка следует пустая ячейка, расположенная
#    по индексу [0, stolbec]
#
#    Сущность <name> также принадлежит сущности <index> и, как правило, содержит в себе значение
#    "Численность обучающихся". Располагается по значению индекса [6, stroka]
#
#    Сущность <unit> содержит в себе сущности:
#        <code>   - [10, stroka]
#        <symbol> - [9,  stroka]
#
#    Значение сущности <deviation> постоянно и равно нулю
#
#    Стандартное смещение: 66 (проценты), 84 (количество человек)
#
#    Необходимо прогонять цикл от 66 до X, чтобы находить последнюю ячейку (после последней ячейки следует пустая ячейка)
#
#    АЛГОРИТМ ГЕНЕРАЦИИ СУЩНОСТЕЙ ТИПА volumeIndex
#
#    1. Находим последнюю строку, начиная с 66
#    2. Запоминаем значение последний строки, чтобы в следующем цикле можно было
#       отталкиваться от этого значения (проверка, не равен ли индекс строки последнему)
#    3. Цикл - while, условие - "пока значение ячейки содержит какое-либо значение"
#    ...
#    N. Создаём коренной элемент		doc.createElement('volumeIndex')
#    						service.appendChild(volumeIndex)
#

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue != "Раздел 1":
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition

# Получение lastPercentElement
firstPercentPosition = firstElementPosition + 13; currentPosition = firstPercentPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue.strip():
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
lastPercentPosition = currentPosition; lastPercentElement = currentPosition - 1

# Получение lastHumanElement
firstHumanPosition = lastPercentElement + 9; currentPosition = firstHumanPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue.strip():
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
lastHumanPosition = currentPosition; lastHumanElement = currentPosition - 1

# Получение количества элементов
numberOfPercentElements = lastPercentPosition - firstPercentPosition
numberOfHumanElements = lastHumanPosition - firstHumanPosition

# Стандартный сдвиг
percentShift = firstHumanPosition

# Элементы
elements = []

# Если количество строк в numberOfPercentElements по какой-то причине не совпадает с numberOfHumanElements
if numberOfPercentElements == numberOfHumanElements:
    totalElements = numberOfPercentElements
else:
    print("Количество строк в numberOfPercentElements не совпадает с numberOfHumanElements")

# Формирование сущностей volumeIndex в цикле for
for i in range(totalElements):

    # Создание новой сущности типа volumeIndex
    volumeIndex = doc.createElement('volumeIndex')
    service.appendChild(volumeIndex)

    # Создание сущности index для последующего её наследования
    index = doc.createElement('index')

    # Создание сущности unit для последующего её наследования
    unit = doc.createElement('unit')

    # Создание сущности deviation (константа)
    deviation = doc.createElement('deviation')
    deviation.appendChild(doc.createTextNode('0'))

    # Создание сущности valueYear для последующего её наследования
    valueYear = doc.createElement('valueYear')

    # Обработка элементов строки
    for currentColumn in range(14):
        currentElement = str(worksheet.cell(percentShift + i, currentColumn).value)
        input = list(worksheet.cell(percentShift + i, 0).value)
        elements.insert(currentColumn, currentElement)

    regNum = doc.createElement('regNum')
    input.insert(7, u'.')
    input.insert(10, u'.')
    input.insert(12, u'.')
    output = "".join(input)
    regNum.appendChild(doc.createTextNode(output))

    name = doc.createElement('name')
    name.appendChild(doc.createTextNode(elements[6]))

    code = doc.createElement('code')
    code.appendChild(doc.createTextNode(str(int(float(elements[10])))))

    symbol = doc.createElement('symbol')
    symbol.appendChild(doc.createTextNode(elements[9]))

    nextYear = doc.createElement('nextYear')
    nextYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[11])), 1)).replace(',', ' ').replace('.', ',')))

    planFirstYear = doc.createElement('planFirstYear')
    planFirstYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[12])), 1)).replace(',', ' ').replace('.', ',')))

    planLastYear = doc.createElement('planLastYear')
    planLastYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[13])), 1)).replace(',', ' ').replace('.', ',')))

#
# Формирование сущностей в порядке, предусмотренном External.xsd
#
#   <index>
#     <regNum></regNum>
#     <name></name>
#     <unit>
#       <code></code>
#       <symbol></symbol>
#     </unit>
#   </index>
#   <deviation></deviation>
#   <valueYear>
#     <nextYear></nextYear>
#     <planFirstYear></planFirstYear>
#     <planLastYear></planLastYear>
#   </valueYear>
#

    volumeIndex.appendChild(index)
    index.appendChild(regNum)
    index.appendChild(name)
    index.appendChild(unit)
    unit.appendChild(code)
    unit.appendChild(symbol)
    volumeIndex.appendChild(deviation)
    volumeIndex.appendChild(valueYear)
    valueYear.appendChild(nextYear)
    valueYear.appendChild(planFirstYear)
    valueYear.appendChild(planLastYear)

# Формирование сущностей indexes в цикле for
for i in range(totalElements):

    # Создание новой сущности типа volumeIndex
    indexes = doc.createElement('indexes')
    service.appendChild(indexes)

    # Обработка элементов строки
    for currentColumn in range(14):
        currentElement = str(worksheet.cell(percentShift + i, currentColumn).value)
        input = list(worksheet.cell(percentShift + i, 0).value)
        elements.insert(currentColumn, currentElement)

    regNum = doc.createElement('regNum')
    input.insert(7, u'.')
    input.insert(10, u'.')
    input.insert(12, u'.')
    output = "".join(input)
    regNum.appendChild(doc.createTextNode(output))
    indexes.appendChild(regNum)

    if str(elements[1]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[1]))
        indexes.appendChild(contentIndex)

    if str(elements[2]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[2]))
        indexes.appendChild(contentIndex)

    if str(elements[3]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[3]))
        indexes.appendChild(contentIndex)

    if str(elements[4]).strip():
        conditionIndex = doc.createElement('conditionIndex')
        conditionIndex.appendChild(doc.createTextNode(elements[4]))
        indexes.appendChild(conditionIndex)

    if str(elements[5]).strip():
        conditionIndex = doc.createElement('conditionIndex')
        conditionIndex.appendChild(doc.createTextNode(elements[5]))
        indexes.appendChild(conditionIndex)

### РАЗДЕЛ №2 ###

# service
service = doc.createElement('service')
ns2_position.appendChild(service)

# service_name
service_name = doc.createElement('name')
service_name.appendChild(doc.createTextNode('Реализация образовательных программ среднего профессионального образования - программ подготовки квалифицированных рабочих, служащих'))
service.appendChild(service_name)

# type_name
type_name = doc.createElement('type')
type_name.appendChild(doc.createTextNode('S'))
service.appendChild(type_name)

# ordinalNumber
ordinalNumber = doc.createElement('ordinalNumber')
ordinalNumber.appendChild(doc.createTextNode('2'))
service.appendChild(ordinalNumber)

# category
category = doc.createElement('category')
service.appendChild(category)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('Физические лица, имеющие основное общее образование'))
category.appendChild(name)

# renderEnactment
renderEnactment = doc.createElement('renderEnactment')
service.appendChild(renderEnactment)

# type
type = doc.createElement('type')
type.appendChild(doc.createTextNode('Закон'))
renderEnactment.appendChild(type)

# author
author = doc.createElement('author')
renderEnactment.appendChild(author)

# fullName
fullName = doc.createElement('fullName')
fullName.appendChild(doc.createTextNode('Федеральный закон'))
author.appendChild(fullName)

# date
date = doc.createElement('date')
date.appendChild(doc.createTextNode('2012-12-29+04:00'))
renderEnactment.appendChild(date)

# number
number = doc.createElement('number')
number.appendChild(doc.createTextNode('273-фз'))
renderEnactment.appendChild(number)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('«Об образовании в Российской Федерации»'))
renderEnactment.appendChild(name)

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue != "Раздел 2":
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition

# Получение lastPercentElement
firstPercentPosition = firstElementPosition + 13; currentPosition = firstPercentPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue.strip():
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
lastPercentPosition = currentPosition; lastPercentElement = currentPosition - 1

# Получение lastHumanElement
firstHumanPosition = lastPercentElement + 9; currentPosition = firstHumanPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue.strip():
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
lastHumanPosition = currentPosition; lastHumanElement = currentPosition - 1

# Получение количества элементов
numberOfPercentElements = lastPercentPosition - firstPercentPosition
numberOfHumanElements = lastHumanPosition - firstHumanPosition

# Стандартный сдвиг
percentShift = firstHumanPosition

# Элементы
elements = []

# Если количество строк в numberOfPercentElements по какой-то причине не совпадает с numberOfHumanElements
if numberOfPercentElements == numberOfHumanElements:
    totalElements = numberOfPercentElements
else:
    print("Количество строк в numberOfPercentElements не совпадает с numberOfHumanElements")

# Формирование сущностей volumeIndex в цикле for
for i in range(totalElements):

    # Создание новой сущности типа volumeIndex
    volumeIndex = doc.createElement('volumeIndex')
    service.appendChild(volumeIndex)

    # Создание сущности для последующего её наследования
    index = doc.createElement('index')

    # Создание сущности unit для последующего её наследования
    unit = doc.createElement('unit')

    # Создание сущности deviation (константа)
    deviation = doc.createElement('deviation')
    deviation.appendChild(doc.createTextNode('0'))

    # Создание сущности valueYear для последующего её наследования
    valueYear = doc.createElement('valueYear')

    # Обработка элементов строки
    for currentColumn in range(14):
        currentElement = str(worksheet.cell(percentShift + i, currentColumn).value)
        input = list(worksheet.cell(percentShift + i, 0).value)
        elements.insert(currentColumn, currentElement)

    regNum = doc.createElement('regNum')
    input.insert(7, u'.')
    input.insert(10, u'.')
    input.insert(12, u'.')
    output = "".join(input)
    regNum.appendChild(doc.createTextNode(output))

    name = doc.createElement('name')
    name.appendChild(doc.createTextNode(elements[6]))

    code = doc.createElement('code')
    code.appendChild(doc.createTextNode(str(int(float(elements[10])))))

    symbol = doc.createElement('symbol')
    symbol.appendChild(doc.createTextNode(elements[9]))

    nextYear = doc.createElement('nextYear')
    nextYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[11])), 1)).replace(',', ' ').replace('.', ',')))

    planFirstYear = doc.createElement('planFirstYear')
    planFirstYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[12])), 1)).replace(',', ' ').replace('.', ',')))

    planLastYear = doc.createElement('planLastYear')
    planLastYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[13])), 1)).replace(',', ' ').replace('.', ',')))

    # Формирование сущностей в порядке, предусмотренном External.xsd
    volumeIndex.appendChild(index)
    index.appendChild(regNum)
    index.appendChild(name)
    index.appendChild(unit)
    unit.appendChild(code)
    unit.appendChild(symbol)
    volumeIndex.appendChild(deviation)
    volumeIndex.appendChild(valueYear)
    valueYear.appendChild(nextYear)
    valueYear.appendChild(planFirstYear)
    valueYear.appendChild(planLastYear)

# Формирование сущностей indexes в цикле for
for i in range(totalElements):

    # Создание новой сущности типа volumeIndex
    indexes = doc.createElement('indexes')
    service.appendChild(indexes)

    # Обработка элементов строки
    for currentColumn in range(14):
        currentElement = str(worksheet.cell(percentShift + i, currentColumn).value)
        input = list(worksheet.cell(percentShift + i, 0).value)
        elements.insert(currentColumn, currentElement)

    regNum = doc.createElement('regNum')
    input.insert(7, u'.')
    input.insert(10, u'.')
    input.insert(12, u'.')
    output = "".join(input)
    regNum.appendChild(doc.createTextNode(output))
    indexes.appendChild(regNum)

    if str(elements[1]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[1]))
        indexes.appendChild(contentIndex)

    if str(elements[2]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[2]))
        indexes.appendChild(contentIndex)

    if str(elements[3]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[3]))
        indexes.appendChild(contentIndex)

    if str(elements[4]).strip():
        conditionIndex = doc.createElement('conditionIndex')
        conditionIndex.appendChild(doc.createTextNode(elements[4]))
        indexes.appendChild(conditionIndex)

    if str(elements[5]).strip():
        conditionIndex = doc.createElement('conditionIndex')
        conditionIndex.appendChild(doc.createTextNode(elements[5]))
        indexes.appendChild(conditionIndex)

### РАЗДЕЛ №3 ###

# service
service = doc.createElement('service')
ns2_position.appendChild(service)

# service_name
service_name = doc.createElement('name')
service_name.appendChild(doc.createTextNode('Реализация основных профессиональных образовательных программ профессионального обучения - программам переподготовки рабочих и служащих'))
service.appendChild(service_name)

# type_name
type_name = doc.createElement('type')
type_name.appendChild(doc.createTextNode('S'))
service.appendChild(type_name)

# ordinalNumber
ordinalNumber = doc.createElement('ordinalNumber')
ordinalNumber.appendChild(doc.createTextNode('3'))
service.appendChild(ordinalNumber)

# category
category = doc.createElement('category')
service.appendChild(category)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('Физические лица, имеющие профессию рабочего или должность служащего'))
category.appendChild(name)

# renderEnactment
renderEnactment = doc.createElement('renderEnactment')
service.appendChild(renderEnactment)

# type
type = doc.createElement('type')
type.appendChild(doc.createTextNode('Закон'))
renderEnactment.appendChild(type)

# author
author = doc.createElement('author')
renderEnactment.appendChild(author)

# fullName
fullName = doc.createElement('fullName')
fullName.appendChild(doc.createTextNode('Федеральный закон'))
author.appendChild(fullName)

# date
date = doc.createElement('date')
date.appendChild(doc.createTextNode('2012-12-29+04:00'))
renderEnactment.appendChild(date)

# number
number = doc.createElement('number')
number.appendChild(doc.createTextNode('273-фз'))
renderEnactment.appendChild(number)

# name
name = doc.createElement('name')
name.appendChild(doc.createTextNode('«Об образовании в Российской Федерации»'))
renderEnactment.appendChild(name)

# Поиск первого элемента
currentPosition = 0; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue != "Раздел 3":
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
firstElementPosition = currentPosition

# Получение lastPercentElement
firstPercentPosition = firstElementPosition + 13; currentPosition = firstPercentPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue.strip():
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
lastPercentPosition = currentPosition; lastPercentElement = currentPosition - 1

# Получение lastHumanElement
firstHumanPosition = lastPercentElement + 9; currentPosition = firstHumanPosition; elementValue = str(worksheet.cell(currentPosition, 0).value)
while elementValue.strip():
    currentPosition += 1; elementValue = str(worksheet.cell(currentPosition, 0).value)
lastHumanPosition = currentPosition; lastHumanElement = currentPosition - 1

# Получение количества элементов
numberOfPercentElements = lastPercentPosition - firstPercentPosition
numberOfHumanElements = lastHumanPosition - firstHumanPosition

# Стандартный сдвиг
percentShift = firstHumanPosition

# Элементы
elements = []

# Если количество строк в numberOfPercentElements по какой-то причине не совпадает с numberOfHumanElements
if numberOfPercentElements == numberOfHumanElements:
    totalElements = numberOfPercentElements
else:
    print("Количество строк в numberOfPercentElements не совпадает с numberOfHumanElements")

# Формирование сущностей volumeIndex в цикле for
for i in range(totalElements):

    # Создание новой сущности типа volumeIndex
    volumeIndex = doc.createElement('volumeIndex')
    service.appendChild(volumeIndex)

    # Создание сущности для последующего её наследования
    index = doc.createElement('index')

    # Создание сущности unit для последующего её наследования
    unit = doc.createElement('unit')

    # Создание сущности deviation (константа)
    deviation = doc.createElement('deviation')
    deviation.appendChild(doc.createTextNode('0'))

    # Создание сущности valueYear для последующего её наследования
    valueYear = doc.createElement('valueYear')

    # Обработка элементов строки
    for currentColumn in range(14):
        currentElement = str(worksheet.cell(percentShift + i, currentColumn).value)
        input = list(worksheet.cell(percentShift + i, 0).value)
        elements.insert(currentColumn, currentElement)

    regNum = doc.createElement('regNum')
    input.insert(7, u'.')
    input.insert(10, u'.')
    input.insert(12, u'.')
    output = "".join(input)
    regNum.appendChild(doc.createTextNode(output))

    name = doc.createElement('name')
    name.appendChild(doc.createTextNode(elements[6]))

    code = doc.createElement('code')
    code.appendChild(doc.createTextNode(str(int(float(elements[10])))))

    symbol = doc.createElement('symbol')
    symbol.appendChild(doc.createTextNode(elements[9]))

    nextYear = doc.createElement('nextYear')
    nextYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[11])), 1)).replace(',', ' ').replace('.', ',')))

    planFirstYear = doc.createElement('planFirstYear')
    planFirstYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[12])), 1)).replace(',', ' ').replace('.', ',')))

    planLastYear = doc.createElement('planLastYear')
    planLastYear.appendChild(doc.createTextNode('{:,}'.format(round(float(str(elements[13])), 1)).replace(',', ' ').replace('.', ',')))


    # Формирование сущностей в порядке, предусмотренном External.xsd
    volumeIndex.appendChild(index)
    index.appendChild(regNum)
    index.appendChild(name)
    index.appendChild(unit)
    unit.appendChild(code)
    unit.appendChild(symbol)
    volumeIndex.appendChild(deviation)
    volumeIndex.appendChild(valueYear)
    valueYear.appendChild(nextYear)
    valueYear.appendChild(planFirstYear)
    valueYear.appendChild(planLastYear)

# Формирование сущностей indexes в цикле for
for i in range(totalElements):

    # Создание новой сущности типа volumeIndex
    indexes = doc.createElement('indexes')
    service.appendChild(indexes)

    # Обработка элементов строки
    for currentColumn in range(14):
        currentElement = str(worksheet.cell(percentShift + i, currentColumn).value)
        input = list(worksheet.cell(percentShift + i, 0).value)
        elements.insert(currentColumn, currentElement)

    regNum = doc.createElement('regNum')
    input.insert(7, u'.')
    input.insert(10, u'.')
    input.insert(12, u'.')
    output = "".join(input)
    regNum.appendChild(doc.createTextNode(output))
    indexes.appendChild(regNum)

    if str(elements[1]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[1]))
        indexes.appendChild(contentIndex)

    if str(elements[2]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[2]))
        indexes.appendChild(contentIndex)

    if str(elements[3]).strip():
        contentIndex = doc.createElement('contentIndex')
        contentIndex.appendChild(doc.createTextNode(elements[3]))
        indexes.appendChild(contentIndex)

    if str(elements[4]).strip():
        conditionIndex = doc.createElement('conditionIndex')
        conditionIndex.appendChild(doc.createTextNode(elements[4]))
        indexes.appendChild(conditionIndex)

    if str(elements[5]).strip():
        conditionIndex = doc.createElement('conditionIndex')
        conditionIndex.appendChild(doc.createTextNode(elements[5]))
        indexes.appendChild(conditionIndex)

#
# ДОПОЛНИТЕЛЬНЫЙ БЛОК ДАННЫХ (ИМПОРТИРОВАНИЕ ЧЕРЕЗ XML НЕ ПРОИЗВОДИТСЯ)
#
# earlyTermination
# earlyTermination = doc.createElement('earlyTermination')
# earlyTermination.appendChild(doc.createTextNode(' '))
# ns2_position.appendChild(earlyTermination)
#
# otherInfo
# otherInfo = doc.createElement('otherInfo')
# otherInfo.appendChild(doc.createTextNode(' '))
# ns2_position.appendChild(otherInfo)
#
# reportRequirements
# reportRequirements = doc.createElement('reportRequirements')
# ns2_position.appendChild(reportRequirements)
#
# deliveryTerm
# deliveryTerm = doc.createElement('deliveryTerm')
# deliveryTerm.appendChild(doc.createTextNode(' '))
# reportRequirements.appendChild(deliveryTerm)
#
# otherRequirement
# otherRequirement = doc.createElement('otherRequirement')
# otherRequirement.appendChild(doc.createTextNode(' '))
# reportRequirements.appendChild(otherRequirement)
#
# otherIndicators
# otherIndicators = doc.createElement('otherIndicators')
# otherIndicators.appendChild(doc.createTextNode(' '))
# reportRequirements.appendChild(otherIndicators)
#
# statementDate
# statementDate = doc.createElement('statementDate')
# statementDate.appendChild(doc.createTextNode('2018-12-29+03:00'))
# ns2_position.appendChild(statementDate)
#
# number
# number = doc.createElement('number')
# number.appendChild(doc.createTextNode('1'))
# ns2_position.appendChild(number)
#
# approverFirstName
# approverFirstName = doc.createElement('approverFirstName')
# approverFirstName.appendChild(doc.createTextNode('Андрей'))
# ns2_position.appendChild(approverFirstName)
#
# approverLastName
# approverLastName = doc.createElement('approverLastName')
# approverLastName.appendChild(doc.createTextNode('Лазарев'))
# ns2_position.appendChild(approverLastName)
#
# approverMiddleName
# approverMiddleName = doc.createElement('approverMiddleName')
# approverMiddleName.appendChild(doc.createTextNode('Александрович'))
# ns2_position.appendChild(approverMiddleName)
#
# approverPosition
# approverPosition = doc.createElement('approverPosition')
# approverPosition.appendChild(doc.createTextNode('Заместитель министра образования Московской области'))
# ns2_position.appendChild(approverPosition)
#

xml_str = doc.toprettyxml(indent="    ")
with open(sys.argv[2], "w") as f:
    f.write(xml_str)
