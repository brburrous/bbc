from dataclasses import dataclass

@dataclass()
class room:
    name: str
    length: float
    width: float
    cleanCode: str
    furnitureCode: str
    protectionCode: str

    def area(self) -> float:
        return self.length * self.width

    