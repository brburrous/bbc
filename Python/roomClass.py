from dataclasses import dataclass

@dataclass()
class room:
    name: str
    length: float = 0
    width: float = 0
    cleanCode: str = ""
    furnitureCode: str = ""
    protectionCode: str = ""
    notes = list()

    def area(self) -> float:
        return self.length * self.width

    def formatList(self) ->list():
        return [self.name, self.length, self.width, self.cleanCode, self.furnitureCode, self.protectionCode]

@dataclass()
class customer:
      name: str
      email: str
      homeNum: int
      workNum: int
      cellNum: int
      otherNum: int
      
@dataclass()
class fileInfo:
    dateStarted: str
    startedBy: str
    dateChanged: str
    changedBy: str
    directions: str
    notes: str

@dataclass()
class customerData:
    JobCustomer: customer
    BillCustomer: customer
    Info: fileInfo
    Rooms = list()


    

