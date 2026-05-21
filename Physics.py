# Физическая часть
import math
# Типы ОЯТ
class SpentNuclearFuel:
    def __init__(self, heat=0, U235=0, U238=0, Pu239=0, RadHazard=0, RadWaste1=0, RadWaste2=0):
        self.HeatGeneration = heat
        self.U235 = U235/100
        self.U238 = U238/100
        self.Pu239 = Pu239/100
        self.Hazard = RadHazard
        self.RadWaste1 = RadWaste1/100
        self.RadWaste2 = RadWaste2/100

FuelTypes = {"ВВЭР-440": SpentNuclearFuel(0, 0, 0, 0),
             "ВВЭР-1000": SpentNuclearFuel(0, 0, 0, 0),
             "ВВЭР-1200": SpentNuclearFuel(0, 0, 0, 0),
             "РБМК-1000": SpentNuclearFuel(0, 0, 0, 0),
             "РБМК-1500": SpentNuclearFuel(0, 0, 0, 0),}
# для BuryReady
FuelClassHeatReady = {"1 класс": 2.0, "2 класс": 0.5, "3 класс": 1.2}

def Heat(n0, t, half_life):
    return n0 / (2 ** (t / half_life))

print(Heat(10, 5, 3))