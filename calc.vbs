dim familysize
dim income
dim poverty
dim message
dim ibr
dim paye
dim ibrmonthly
dim payemonthly


familysize = Inputbox("What is your household family size?","Enter here")
income = Inputbox("What is your household income. Note: If you are single (or if you file your taxes single) only put YOUR income.","Enter here")

'calculation for 150% poverty guidelines
poverty = 17655 + ((familysize - 1) * 6240)

'calculation for IBR payment (15% discretionary income)
ibr = ((income - poverty) * 0.15) / 12

'calculation for PAYE (10% discretionary income)
paye = ((income - poverty) * 0.10) / 12

if ibr < 0 then
ibr = 0
end if

if paye < 0 then
paye = 0
end if

ibrmonthly = Round(ibr,2)
payemonthly = Round(paye,2)



message = MsgBox("Note: this is not accurate if you live in Alaska or Hawaii" & vbNewLine & vbNewLine & "PAYE estimated payment:" & " " & "$" & payemonthly & vbNewLine & "IBR estimated payment:" & " " & "$" & ibrmonthly, 0)
 