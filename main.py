from xlrd import open_workbook
from sqlalchemy import *

class Item:
	def __init__(self):
		self.uniqueId = ''
		self.theOtherColumn = False
		self.genesse = False
		self.shiawassee = False
		self.lapeer = False
		self.stClair = False
		self.tuscola = False
		self.huron = False
		self.sanilac = False
		self.statewide = False
		self.name = ''
		self.parentOrganization = ''
		self.xLong = 0.0
		self.yLat = 0.0
		self.address = ''
		self.address2 = ''
		self.city = ''
		self.zipCode = ''
		self.zipCode2 = ''
		self.state = ''
		self.countyOfOrigin = ''
		self.phoneNumber = ''
		self.mainWebsite = ''
		self.website = ''
		self.contact = ''
		self.email = ''
		self.services = ''
		self.description = ''
		self.restriction = ''
		self.phase = '' 
		self.any = False				#phase 0
		self.idea = False				#phase 1
		self.startUp = False 			#phase 2 
		self.existingBusiness = False  	#phase 3
		self.notes = ''
		self.addressFormat = ''
		self.businessAssistance = False
		self.networking = False
		self.rnd = False
		self.funding = False
		self.office = False
		self.markerspace = False
		self.kitchen = False
		self.license = False
	#setters and getters
	def setUniqueId(self,uniqueId):
		self.uniqueId = uniqueId if uniqueId is not None else ''
	def getUniqueId(self):
		return str(self.uniqueId)
	def setTheOtherColumn(self,theOtherColumn):
		self.theOtherColumn = True if theOtherColumn is 'X' else False
	def getTheOtherColumn(self):
		return self.theOtherColumn
	def setGenesse(self,genesse):
		self.genesse = True if genesse is 'X' else False
	def getGenesse(self):
		return self.genesse
	def setShiawassee(self,shiawassee):
		self.shiawassee = True if shiawassee is 'X' else False
	def getShiawassee(self):
		return self.shiawassee
	def setLapeer(self,lapeer):
		self.lapeer = True if lapeer is 'X' else False
	def getLapeer(self):
		return self.lapeer
	def setStClair(self,stClair):
		self.stClair = True if stClair is 'X' else False
	def getStClair(self):
		return self.stClair
	def setTuscola(self,tuscola):
		self.tuscola = True if tuscola is 'X' else False
	def getTuscola(self):
		return self.tuscola
	def setHuron(self,huron):
		self.huron = True if huron is 'X' else False
	def getHuron(self):
		return self.huron
	def setSanilac(self,sanilac):
		self.sanilac = True if sanilac is 'X' else False
	def getSanilac(self):
		return self.sanilac
	def setStatewide(self,statewide):
		self.statewide = True if statewide is 'X' else False
	def getStatewide(self):
		return self.statewide
	def setName(self,name):
		self.name = name if name is not None else ''
	def getName(self):
		return str(self.name)
	def setParentOrganization(self,parentOrganization):
		self.parentOrganization = parentOrganization if parentOrganization is not None else ''
	def getParentOrganization(self):
		return str(self.parentOrganization)
	def setXLong(self,xLong):
		self.xLong = xLong if xLong is not '' else 0.0
	def getXLong(self):
		return self.xLong
	def setYLat(self,yLat):
		self.yLat = yLat if yLat is not '' else 0.0
	def getYLat(self):
		return self.yLat
	def setAddress(self,address):
		self.address = address if address is not None else ''
	def getAddress(self):
		return str(self.address)
	def setAddress2(self,address2):
		self.address2 = address2 if address2 is not None else ''
	def getAddress2(self):
		return str(self.address2)
	def setCity(self,city):
		self.city = city if city is not None else ''
	def getCity(self):
		return str(self.city)
	def setZipCode(self,zipCode):
		self.zipCode = zipCode if zipCode is not '' else ''
	def getZipCode(self):
		return str(self.zipCode)
	def setZipCode2(self,zipCode2): 
		self.zipCode2 = zipCode2 if zipCode2 is not '' else ''
	def getZipCode2(self):
		return str(self.zipCode2)
	def setState(self,state):
		self.state = state if state is not None else ''
	def getState(self):
		return str(self.state)
	def setCountyOfOrigin(self,countyOfOrigin):
		self.countyOfOrigin = countyOfOrigin if countyOfOrigin is not None else ''
	def getCountyOfOrigin(self):
		return str(self.countyOfOrigin)
	def setPhoneNumber(self,phoneNumber):
		self.phoneNumber = phoneNumber if phoneNumber is not None else ''
	def getPhoneNumber(self):
		return str(self.phoneNumber)
	def setMainWebsite(self,mainWebsite):
		self.mainWebsite = mainWebsite if mainWebsite is not None else ''
	def getMainWebsite(self):
		return str(self.mainWebsite)
	def setWebsite(self,website):
		self.website = website if website is not None else ''
	def getWebsite(self):
		return str(self.website)
	def setContact(self,contact):
		self.contact = contact
	def getContact(self):
		return str(self.contact)
	def setEmail(self,email):
		self.email = email if email is not None else ''
	def getEmail(self):
		return str(self.email)
	def setServices(self,services):
		self.services = services if services is not None else ''
	def getServices(self):
		return str(self.services)
	def setDescription(self,description):
		self.description = description if description is not None else ''
	def getDescription(self):
		return str(self.description)
	def setRestriction(self,restriction):
		self.restriction = restriction if restriction is not None else ''
	def getRestriction(self):
		return str(self.restriction)
	def setPhase(self,phase):
		self.phase = phase if phase is not None else ''
	def getPhase(self):
		return str(self.phase)
	def setAny(self,any):
		self.any = True if any is 'X' else False
	def getAny(self):
		return self.any
	def setIdea(self,idea):
		self.idea = True if idea is 'X' else False
	def getIdea(self):
		return self.idea
	def setStartUp(self,startUp):
		self.startUp = True if startUp is 'X' else False
	def getStartUp(self):
		return self.startUp
	def setExistingBusiness(self,existingBusiness):
		self.existingBusiness = True if existingBusiness is 'X' else False
	def getExistingBusiness(self):
		return self.existingBusiness
	def setNotes(self,notes):
		self.notes = notes if notes is not None else ''
	def getNotes(self):
		return str(self.notes)
	def setAddressFormat(self,addressFormat):
		self.addressFormat = addressFormat if addressFormat is not None else ''
	def getAddressFormat(self):
		return str(self.addressFormat)
	def setBusinessAssistance(self,businessAssistance):
		self.businessAssistance = True if businessAssistance is 'X' else False
	def getBusinessAssistance(self):
		return self.businessAssistance
	def setNetworking(self,networking):
		self.networking = True if networking is 'X' else False
	def getNetworking(self):
		return self.networking
	def setRnd(self,rnd):
		self.rnd = True if rnd is 'X' else False
	def getRnd(self):
		return self.rnd
	def setFunding(self,funding):
		self.funding = True if funding is 'X' else False
	def getFunding(self):
		return self.funding
	def setOffice(self,office):
		self.office = True if office is 'X' else False
	def getOffice(self):
		return self.office
	def setMarkerspace(self,markerspace):
		self.markerspace = True if markerspace is 'X' else False
	def getMarkerspace(self):
		return self.markerspace
	def setKitchen(self,kitchen):
		self.kitchen = True if kitchen is 'X' else False
	def getKitchen(self):
		return self.kitchen
	def setLicense(self,license):
		self.license = True if license is 'X' else False
	def getLicense(self):
		return self.license
	#overrides
	def __str__(self):
		return self.getPhoneNumber() + '\n'
	def __eq__(self,other):
		return self.name == other.name

b = open_workbook('data/data2.xlsx')
s = b.sheet_by_name('Linear')
items = []
for row in range(1,s.nrows):
	i = Item()
	i.setUniqueId(s.cell(row,0).value)
	i.setTheOtherColumn(s.cell(row,1).value)
	i.setGenesse(s.cell(row,2).value)
	i.setShiawassee(s.cell(row,3).value)
	i.setLapeer(s.cell(row,4).value)
	i.setStClair(s.cell(row,5).value)
	i.setTuscola(s.cell(row,6).value)
	i.setHuron(s.cell(row,7).value)
	i.setSanilac(s.cell(row,8).value)
	i.setStatewide(s.cell(row,9).value)
	i.setName(s.cell(row,10).value)
	i.setParentOrganization(s.cell(row,11).value)
	i.setXLong(s.cell(row,12).value)
	i.setYLat(s.cell(row,13).value)
	i.setAddress(s.cell(row,14).value)
	i.setAddress2(s.cell(row,15).value)
	i.setCity(s.cell(row,16).value)
	i.setZipCode(s.cell(row,17).value)
	i.setZipCode2(s.cell(row,18).value)
	i.setState(s.cell(row,19).value)
	i.setCountyOfOrigin(s.cell(row,20).value)
	i.setPhoneNumber(s.cell(row,21).value)
	i.setMainWebsite(s.cell(row,22).value)
	i.setWebsite(s.cell(row,23).value)
	i.setContact(s.cell(row,24).value)
	i.setEmail(s.cell(row,25).value)
	i.setServices(s.cell(row,26).value)
	i.setDescription(s.cell(row,27).value)
	i.setRestriction(s.cell(row,28).value)
	i.setPhase(s.cell(row,29).value)
	i.setAny(s.cell(row,30).value)
	i.setIdea(s.cell(row,31).value)
	i.setStartUp(s.cell(row,32).value)
	i.setExistingBusiness(s.cell(row,33).value)
	i.setNotes(s.cell(row,34).value)
	i.setAddressFormat(s.cell(row,35).value)
	i.setBusinessAssistance(s.cell(row,36).value)
	i.setNetworking(s.cell(row,37).value)
	i.setRnd(s.cell(row,38).value)
	i.setFunding(s.cell(row,39).value)
	i.setOffice(s.cell(row,40).value)
	i.setMarkerspace(s.cell(row,41).value)
	i.setKitchen(s.cell(row,42).value)
	i.setLicense(s.cell(row,43).value)
	items.append(i)

# freq = {}
# for i in items:
# 	if i.getName() not in freq:
# 		freq[i.getName()] = 1
# 	else:
# 		freq[i.getName()] += 1

# for key,value in freq.items():
# 	if value > 1:
# 		print(key,value)


engine = create_engine('mysql+pymysql://root:vmdb@localhost:3306/gis',encoding='utf/8')
connection = engine.connect()
for item in items:
	connection.execute("insert into info (uniqueId,theOtherColumn,genesee,shiawassee,lapeer,stClair,tuscola,huron,sanilac,statewide,name,parentOrganization,xLong,yLat,address,address2,city,zipCode,zipCode2,state,countyOfOrigin,phoneNumber,mainWebsite,website,contact,email,services,description,restriction,phase,any,idea,startUp,existingBusiness,notes,addressFormat,businessAssistance,networking,rnd,funding,office,makerspace,kitchen,license) values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);",
		(item.getUniqueId(),item.getTheOtherColumn(),item.getGenesse(),item.getShiawassee(),
			item.getLapeer(),item.getStClair(),item.getTuscola(),item.getHuron(),item.getSanilac(),
			item.getStatewide(),item.getName(),item.getParentOrganization(),item.getXLong(),
			item.getYLat(),item.getAddress(),item.getAddress2(),item.getCity(),item.getZipCode(),
			item.getZipCode2(),item.getState(),item.getCountyOfOrigin(),item.getPhoneNumber(),
			item.getMainWebsite(),item.getWebsite(),item.getContact(),item.getEmail(),
			item.getServices(),item.getDescription(),item.getRestriction(),item.getPhase(),
			item.getAny(),item.getIdea(),item.getStartUp(),item.getExistingBusiness(),
			item.getNotes(),item.getAddressFormat(),
			item.getBusinessAssistance(),item.getNetworking(),item.getRnd(),item.getFunding(),
			item.getOffice(),item.getMarkerspace(),item.getKitchen(),item.getLicense()))
	print(item.getName(),'added')

connection.close()