from dataclasses import dataclass

@dataclass()
class room:
    name: str
    length: float = 0
    width: float = 0
    cleanCode: str = ""
    furnitureCode: str = ""
    protectionCode: str = ""

    def area(self) -> float:
        return self.length * self.width

    def formatList(self) ->list():
        return [self.name, self.length, self.width, self.cleanCode, self.furnitureCode, self.protectionCode]