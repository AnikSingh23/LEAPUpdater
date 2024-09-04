from win32com.client import Dispatch

L = Dispatch('LEAP.LEAPApplication')
B = L.Branch("\Key Assumptions\TotalElec")
V = B.Variables("Activity Level")

L.ActiveArea = "Test Bed"
L.ActiveScenario = "Current Accounts"
print(V.Expression)
C.Expression = "Interp(ab\\ab com.xlsx,table 1!c10:ad10, table 1!c14:ad14)"
