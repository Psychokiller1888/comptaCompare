import re
from pathlib import Path

from PyInquirer import prompt
from openpyxl import load_workbook

MATCH_DATE = re.compile(r'^[\d]{2}\.[\d]{2}\.[\d]{4}$')
#MATCH_RAIFFEISEN = re.compile(r"(?P<date>[\d]{2}\.[\d]{2}\.[\d]{2}).*?(?P<type>Crédit|Ordre|Système de recouvrement direct|Versement).*?(?P<amount>[0-9']+\.[0-9]{2}) (?P<solde>[0-9']+\.[0-9]{2})$")
#MATCH_ABACUS = re.compile(r"(?P<date>[\d]{2}\.[\d]{2}\.[\d]{4}).*?(?P<amount>[0-9']+\.[0-9]{2}) (?P<solde>[0-9']+\.[0-9]{2})")
#MATCH_ABACUS_START_SOLDE = re.compile(r"Solde y\.c\. report (?P<startSolde>[\d']+\.[\d]{2})")
#MATCH_RAIFFEISEN_START_SOLDE = re.compile(r"Report de solde (?P<startSolde>[\d']+\.[\d]{2})")
#MATCH_RAIFFEISEN_END_SOLDE = re.compile(r"Solde en votre faveur.*?(?P<endSolde>[\d']+\.[\d]{2})")

abacusExportDir = Path('abacusExports')
bankExportDir = Path('bankExports')


def parseMoney(string: str) -> float:
	try:
		return float(string.replace("'", '').strip())
	except:
		raise


def isFloat(string: str) -> bool:
	try:
		float(string)
		return True
	except:
		return False


abacusExports = abacusExportDir.glob('*.xlsx')
bankExports = bankExportDir.glob('*.xlsx')

questions = [
	{
		'type': 'list',
		'name': 'abacusFile',
		'message': 'Select abacus account export',
		'choices': [file.stem for file in abacusExports]
	},
	{
		'type': 'list',
		'name': 'bankFile',
		'message': 'Select bank account export',
		'choices': [file.stem for file in bankExports]
	},
	{
		'type': 'confirm',
		'name': 'knownDifferenceQuestion',
		'message': 'Are there any known difference with the previous month that you cannot currently correct?',
		'default': False
	},
	{
		'type': 'input',
		'name': 'knownDifference',
		'message': 'Please input the difference to apply on the Abacus end saldo of the previous month',
		'validate': lambda val: isFloat(val),
		'when': lambda ans: 'knownDifferenceQuestion' in ans and ans['knownDifferenceQuestion']
	}
]

answers = prompt(questions=questions)

if not answers:
	exit(1)

abacusFile = open(Path(abacusExportDir, f'{answers["abacusFile"]}.xlsx'), mode='rb')

print('\nAnalyzing Abacus...')
workbook = load_workbook(filename=abacusFile)
endSaldoAbacus = 0
abacusCredits = dict()
abacusDebits = dict()
linesAbacus = 0
sheet = workbook.active

startSaldoAbacus = sheet['I4'].value

prevSaldo = 0
for row in sheet.rows:
	row0 = str(row[0].value)
	if MATCH_DATE.match(row0):
		date = row0
		if row[6].value is not None:
			abacusDebits.setdefault(date, list())
			abacusDebits[date].append(float(row[6].value))
			linesAbacus += 1
			prevSaldo = row[8].value
		elif row[7].value is not None:
			abacusCredits.setdefault(date, list())
			abacusCredits[date].append(float(row[7].value))
			linesAbacus += 1
			prevSaldo = row[8].value

	elif row0.startswith('Solde') and row0 != 'Solde y.c. report':
		endSaldoAbacus = prevSaldo
		break

print('\nAnalyzing Raiffeisen...')
bankCredits = dict()
bankDebits = dict()
endSaldoRaiffeisen = 0
linesRaiffeisen = 0
bankFile = open(Path(bankExportDir, f'{answers["bankFile"]}.xlsx'), mode='rb')
workbook = load_workbook(filename=bankFile)
sheet = workbook.active

startSaldoRaiffeisen = sheet['E2'].value - sheet['D2'].value if sheet['E2'].value > 0 else sheet['E2'].value + sheet['D2'].value
for row in sheet.rows:
	if row[0].value.lower() == 'iban':
		continue

	date = row[1].value

	try:
		date = date.strftime('%d.%m.%y')
	except:
		continue

	if row[3].value == 0:
		continue
	elif row[3].value < 1:
		bankDebits.setdefault(date, list())
		bankDebits[date].append(abs(float(row[3].value)))
		linesRaiffeisen += 1
	else:
		bankCredits.setdefault(date, list())
		bankCredits[date].append(float(row[3].value))
		linesRaiffeisen += 1

	endSaldoRaiffeisen = float(row[4].value)

print('\nDocument analyze done!')

if linesRaiffeisen > linesAbacus:
	print(f'\n- Raiffeisen has MORE entries than Abacus')
elif linesRaiffeisen < linesAbacus:
	print(f'\n- Raiffeisen has LESS entries than Abacus')

print(f'-- Raiffeisen lines: {linesRaiffeisen}')
print(f'-- Abacus lines: {linesAbacus}')

if startSaldoRaiffeisen != startSaldoAbacus:
	if 'knownDifference' in answers and float(answers['knownDifference']) != 0:
		print(f'\n- Starting saldo are not the same, applying known difference to Abacus start saldo')
		startSaldoAbacus += float(answers['knownDifference'])
		if startSaldoRaiffeisen != startSaldoAbacus:
			print(f'-- Even after applying {answers["knownDifference"]} to Abacus start saldo, saldo are not matching, did you correct the month before?')
		else:
			print(f'-- After applying {answers["knownDifference"]} to Abacus start saldo, the account start saldo match!')
	else:
		print(f'\n- Starting saldo are not the same, did you already correct the month before?')
else:
	print(f'\n- Starting saldo are matching, previous months are ok!')

print(f'-- Raiffeisen start saldo: {startSaldoRaiffeisen}')
print(f'-- Abacus start saldo: {startSaldoAbacus}')

print(f'\n-- Raiffeisen end saldo: {endSaldoRaiffeisen}')
print(f'-- Abacus end saldo: {endSaldoAbacus}')

if endSaldoRaiffeisen != endSaldoAbacus:
	if 'knownDifference' in answers and float(answers['knownDifference']) != 0:
		print(f'- End saldo are not the same, applying known difference to Abacus end saldo')
		endSaldoAbacus += float(answers['knownDifference'])
		if endSaldoRaiffeisen != endSaldoAbacus:
			print(f'- Even after applying {answers["knownDifference"]} to Abacus end saldo, saldo are not matching, you got work to do!')
		else:
			print(f'- After applying {answers["knownDifference"]} to Abacus end saldo, the account saldo match!')
	else:
		print(f'\n- Ending saldo are not the same, you have work to do!')
else:
	print(f'\n- Ending saldo are matching, so far so good!')



print('\nMatching Raiffeisen credits to Abacus debits...')

notFoundBankCredits = list()
potentialEndSaldoCorrection = 0
for date, amounts in bankCredits.copy().items():
	for amount in amounts:
		found = False
		for abacusDate, abacusAmounts in abacusDebits.copy().items():
			for abacusAmount in abacusAmounts.copy():
				if amount == abacusAmount:
					found = True
					abacusDebits[abacusDate].remove(abacusAmount)
					break
			if found:
				break
		if not found:
			print(f'- Crédit Raiffeisen de {amount} CHF du {date} non trouvé dans les écritures Abacus')
			potentialEndSaldoCorrection += amount
			notFoundBankCredits.append(amount)

if potentialEndSaldoCorrection > 0:
	print(f'\n-- Applying potential correction, new end saldo on Abacus: {round(endSaldoAbacus + potentialEndSaldoCorrection)}')


print('\nMatching Raiffeisen debits to Abacus credits...')
notFoundBankDebits = list()
for date, amounts in bankDebits.copy().items():
	for amount in amounts:
		found = False
		for abacusDate, abacusAmounts in abacusCredits.copy().items():
			for abacusAmount in abacusAmounts.copy():
				if amount == abacusAmount:
					found = True
					abacusCredits[abacusDate].remove(abacusAmount)
					break
			if found:
				break
		if not found:
			print(f'- Débit Raiffeisen de {amount} CHF du {date} non trouvé dans les écritures Abacus')
			potentialEndSaldoCorrection -= amount
			notFoundBankDebits.append(amount)

newPotentialAbacusSaldo = endSaldoAbacus
if potentialEndSaldoCorrection != 0:
	newPotentialAbacusSaldo = round(endSaldoAbacus + potentialEndSaldoCorrection, 2)
	print(f'\n-- Applying potential correction, new end saldo on Abacus: {newPotentialAbacusSaldo}')
	if newPotentialAbacusSaldo == endSaldoRaiffeisen:
		print('--- Correcting missing parts would fix the account and bring it to same level as the bank!')
	else:
		print("--- Applying corrections does not fix the account, there's something else...")

print('\nThese entries are not matched to Raiffeisen entries')
correction = 0
inverseAbacusCredits = list()
inverseAbacusDebits = list()
removeFromAbacus = list()
print('- Credits')
for date, listing in abacusCredits.items():
	for amount in listing:
		print(f'-- {date} {amount}')
		correction += amount
		if amount in notFoundBankCredits:
			inverseAbacusCredits.append(str(amount))
		else:
			removeFromAbacus.append(str(amount))

print('- Debits')
for date, listing in abacusDebits.items():
	for amount in listing:
		print(f'-- {date} {amount}')
		correction -= amount
		if amount in notFoundBankDebits:
			inverseAbacusDebits.append(str(amount))
		else:
			removeFromAbacus.append(str(amount))

if correction != 0:
	print('\n-- Adding non matched credits and removing non matched debits...')
	result = round(newPotentialAbacusSaldo + correction, 2)
	print(f'--- Calculated result on Abacus account: {result} {f"(should be {endSaldoRaiffeisen})" if result != endSaldoRaiffeisen else ""}')
	if result == endSaldoRaiffeisen:
		if inverseAbacusCredits:
			print(f'---- Inverse the following entrie(s) to match the accounts: {", ".join(inverseAbacusCredits)}')

		if inverseAbacusDebits:
			print(f'---- Inverse the following entrie(s) to match the accounts: {", ".join(inverseAbacusDebits)}')

		if removeFromAbacus:
			print(f'---- Remove the following entrie(s) to match the accounts: {", ".join(removeFromAbacus)}')
	else:
		print('--- Accounts not matching, something else is wrong...')
