DEFINT A-Z
	DECLARE SUB ClearSheetAfterMenu ()
	DECLARE SUB GetText ()
	DECLARE SUB GetFormula ()
	DECLARE SUB GetNumber ()
	DECLARE SUB GetEnd ()
	DECLARE SUB ChangeDir ()
	DECLARE SUB SaveSpread ()
	DECLARE SUB OpenSpread ()
	DECLARE SUB SetWidths ()
	DECLARE SUB SetFile ()
	DECLARE SUB ShowMenu (MenuNumber%)
	DECLARE SUB Delete ()
	DECLARE SUB Insert ()
	DECLARE SUB InsertRow ()
	DECLARE SUB InsertCol ()
	DECLARE SUB DeleteRow ()
	DECLARE SUB DeleteCol ()
	DECLARE SUB ShowLetters (Letter$)
	DECLARE SUB ShowNumbers (bgin AS INTEGER)
	DECLARE SUB EraseCursor ()
	DECLARE SUB ClearSheet ()
	DECLARE SUB GetEndOfLine (col%, row%, Value$)
	DECLARE SUB CheckInFormula ()
	DECLARE SUB NestedForm (In$, Out$)
	DECLARE SUB FindRange (In$, Out$)
	DECLARE SUB CW ()
	DECLARE SUB SelectFormat ()
	DECLARE SUB NewSpread ()
	DECLARE SUB ReCalc ()
	DECLARE SUB CalcForm (col%, row%)
	DECLARE SUB SplitForm (st$, newst$(), s)
	DECLARE SUB ChangeFormat ()
	DECLARE SUB Translate (Start$, end$)
	DECLARE SUB FileGet (file$)
	DECLARE SUB DOS ()
	DECLARE SUB ShowData (bgin%, Letter$)
	DECLARE SUB Box (Row1%, Col1%, Row2%, Col2%, BoxStr AS STRING)
	DECLARE SUB GetCurrStuff ()
	DECLARE SUB GetFreeSpace ()
	DECLARE SUB SaveINI ()
	DECLARE SUB LoadINI ()
	DECLARE SUB ChooseColor (colr%)
	DECLARE SUB DrawInfo ()
	DECLARE SUB CalcHalves ()
	DECLARE SUB ShowExtra ()
	DECLARE SUB ClearBuffer ()
	DECLARE SUB CheckKeys ()
	DECLARE SUB ShowStat (stat$)
	DECLARE SUB ShowCurCell ()
	DECLARE SUB N2S (Num$, st$)
	DECLARE SUB ClearTop ()
	DECLARE SUB DrawForm ()
	DECLARE SUB Intro ()
	DECLARE SUB MainSystem ()
	DECLARE SUB Center (row%, text$)
	DECLARE SUB SpinThing (row%, col%, spins%)
	DECLARE FUNCTION MMenu% (CurrChoiceX, MaxChoice, choice$(), ItemRow(), ItemCol(), Hot$, BarMode)
	DECLARE FUNCTION GetString$ (row%, col%, Start$, end$, Vis%, Max%)
	DECLARE FUNCTION GetNumeric$ (row%, col%, Start$, end$, Vis%, Max%)

	TYPE CellData
		CellValue AS STRING * 100
		Usg AS STRING * 21
		Formula AS STRING * 100
		ID AS STRING * 2
	END TYPE

	CONST TRUE = -1
	CONST FALSE = NOT TRUE
	CONST MaxX = 10
	CONST MaxY = 25
	DIM SHARED ShowCell
	DIM SHARED DefaultDir$
		DIM SHARED W(MaxX) AS INTEGER
		DIM SHARED Cell(MaxX, MaxY) AS CellData
	DIM SHARED CalcMode AS STRING * 10
	DIM SHARED Iter AS STRING * 3
	DIM SHARED Order AS STRING * 11
	DIM SHARED MainCurr AS INTEGER
	DIM SHARED MC AS INTEGER
	DIM SHARED Changed AS INTEGER
	DIM SHARED SubMenu AS INTEGER
	DIM SHARED MenuRow AS INTEGER
	DIM SHARED MenuLeft AS INTEGER
	DIM SHARED MenuRight AS INTEGER
	DIM SHARED Clean AS INTEGER
	DIM SHARED Check AS INTEGER
	DIM SHARED LastChar AS STRING * 1
	DIM SHARED CellRow AS INTEGER
	DIM SHARED CellCol AS INTEGER
	DIM SHARED FirstX AS INTEGER
	DIM SHARED FirstY AS INTEGER
	DIM SHARED CurrCell AS STRING * 1
	DIM SHARED currRow AS INTEGER
	DIM SHARED MenuHotKey AS INTEGER
	DIM SHARED MenuHighlight AS INTEGER
	DIM SHARED MenuRegular AS INTEGER
	DIM SHARED SheetColor AS INTEGER
	DIM SHARED StatColor AS INTEGER
	DIM SHARED InfoColor AS INTEGER
	DIM SHARED CursorColor AS INTEGER
	DIM SHARED DefaultFile AS STRING * 12
	DIM SHARED FreeSpace AS STRING * 13
	DIM SHARED CurrDrive AS STRING * 1
	DIM SHARED CurrDir AS STRING * 128
	DIM SHARED CurrSheet AS STRING * 12
	DIM SHARED NumType AS STRING * 2
	DIM SHARED LetterMid AS INTEGER
	DIM SHARED Decimal AS STRING * 1
	DIM SHARED Move$
	DIM SHARED TT AS INTEGER
	DIM SHARED CellData$
	DIM SHARED DblBox AS STRING * 8
	DIM SHARED SngBox AS STRING * 8

	ON ERROR GOTO ErrorTrap

CellRow = 4
CellCol = 7
DblBox = "*ͻ***Ⱥ"
SngBox = "*Ŀ*****"
CalcMode = "Background"
Iter = "1"
Order = "Natural"
NumType = "P2"
Decimal = "2"
CurrCell = "A"
currRow = 1
MenuHighlight = 7
MenuRegular = 7
MenuHotKey = 15
SheetColor = 15
StatColor = 15
InfoColor = 15
CursorColor = 7
CurrDrive = " "
Check = 1
Intro
MainSystem
SYSTEM

ErrorTrap:
RESUME NEXT

SUB Box (Row1, Col1, Row2, Col2, BoxStr AS STRING)
	BoxWidth = Col2 - Col1 + 1

	LOCATE Row1, Col1
	PRINT LEFT$(BoxStr, 1); STRING$(BoxWidth - 2, MID$(BoxStr, 2, 1)); MID$(BoxStr, 3, 1);

	FOR a = Row1 + 1 TO Row2 - 1
		LOCATE a, Col1
		PRINT MID$(BoxStr, 8, 1); SPACE$(BoxWidth - 2); MID$(BoxStr, 4, 1);
	NEXT a

	LOCATE Row2, Col1
	PRINT MID$(BoxStr, 7, 1); STRING$(BoxWidth - 2, MID$(BoxStr, 6, 1)); MID$(BoxStr, 5, 1);
END SUB

SUB CalcForm (col, row)
	DIM Num$(100), Oper$(100), Letter$(100)
		IF RTRIM$(Cell(col, row).Formula) = "" THEN EXIT SUB

		NewForm$ = Cell(col, row).Formula
FOR Itt = 1 TO VAL(Iter)

		Range = 0
		FOR t = 1 TO LEN(NewForm$)
			IF MID$(NewForm$, t, 1) = "[" THEN
				Range = Range + 1
			END IF
		NEXT
		IF Range > 0 THEN
			FOR t = 1 TO Range
				FindRange NewForm$, NewForm$
			NEXT
		END IF

		Nested = 0
		FOR t = 1 TO LEN(NewForm$)
			IF MID$(NewForm$, t, 1) = "(" THEN
				Nested = Nested + 1
			END IF
		NEXT
		IF Nested > 0 THEN
			FOR t = 1 TO Nested
				NestedForm NewForm$, NewForm$
			NEXT
		END IF

				SplitForm NewForm$, Num$(), Num

		FOR t = 1 TO Num
				Letter$(t) = LEFT$(Num$(t), 1)
					IF Letter$(t) = "9" OR Letter$(t) = "8" OR Letter$(t) = "7" OR Letter$(t) = "6" OR Letter$(t) = "5" OR Letter$(t) = "4" OR Letter$(t) = "3" OR Letter$(t) = "2" OR Letter$(t) = "1" OR Letter$(t) = "0" THEN Letter$(t) = " "
			IF RIGHT$(Num$(t), 1) = "+" OR RIGHT$(Num$(t), 1) = "-" OR RIGHT$(Num$(t), 1) = "/" OR RIGHT$(Num$(t), 1) = "*" OR RIGHT$(Num$(t), 1) = "S" THEN
				Oper$(t) = RIGHT$(Num$(t), 1)
				IF Letter$(t) <> " " THEN
					Num$(t) = LEFT$(Num$(t), LEN(Num$(t)) - 1)
				ELSE
					Num$(t) = LEFT$(Num$(t), LEN(Num$(t)))
				END IF
			ELSE
				Oper$(t) = "+"
			END IF
				Num$(t) = RIGHT$(Num$(t), LEN(Num$(t)) - 1)
		NEXT
				lt = (ASC(Letter$(1)) - 64)
				dt = (VAL(Num$(1)))
				IF Letter$(1) <> " " THEN
					Nxt# = VAL(Cell(lt, dt).CellValue)
				ELSE
					Nxt# = VAL(Num$(1))
				END IF
					TotalValue# = Nxt#
		FOR t = 2 TO Num
				lt = (ASC(Letter$(t)) - 64)
				dt = (VAL(Num$(t)))
				IF Letter$(t) <> " " THEN
					Nxt# = VAL(Cell(lt, dt).CellValue)
				ELSE
					Nxt# = VAL(Num$(t))
				END IF

					IF t = 1 THEN TotalValue# = Nxt#
			IF Oper$(t - 1) = "^" THEN TotalValue# = TotalValue# ^ Nxt#
			IF Oper$(t - 1) = "*" THEN TotalValue# = TotalValue# * Nxt#
			IF Oper$(t - 1) = "/" THEN TotalValue# = TotalValue# / Nxt#
			IF Oper$(t - 1) = "+" THEN TotalValue# = TotalValue# + Nxt#
			IF Oper$(t - 1) = "-" THEN TotalValue# = TotalValue# - Nxt#
		NEXT
NEXT

		TotalValue$ = STR$(TotalValue#)
		TotalValue$ = RTRIM$(TotalValue$)
		TotalValue$ = LTRIM$(TotalValue$)
				IF Num = 1 THEN
					Cell(col, row).CellValue = Cell((ASC(Letter$(1)) - 64), VAL(Num$(1))).CellValue
					Cell(col, row).Usg = Cell((ASC(Letter$(1)) - 64), VAL(Num$(1))).Usg
					Cell(col, row).ID = Cell((ASC(Letter$(1)) - 64), VAL(Num$(1))).ID
					EXIT SUB
				END IF

				Cell(col, row).CellValue = TotalValue$
						IF LEN(Cell(col, row).CellValue) > 8 AND LEFT$(Cell(col, row).ID, 1) = "P" AND LEFT$(NumType, 1) <> "C" THEN
							rt$ = NumType
							NumType = "S" + Decimal
							Translate RTRIM$(Cell(col, row).CellValue), NewUs$
							Cell(col, row).Usg = NewUs$
							Cell(col, row).ID = NumType
							NumType = rt$
						ELSE
							rt$ = NumType
							NumType = Cell(col, row).ID
								IF LEFT$(NumType, 1) = "\" THEN
									NumType = rt$
									Cell(col, row).ID = NumType
								END IF
								Translate RTRIM$(Cell(col, row).CellValue), NewUs$
								Cell(col, row).Usg = NewUs$
							NumType = rt$
						END IF

END SUB

SUB CalcHalves
	ShowStat "WAIT"
	FOR t = 1 TO MaxX
		W(t) = 9
	NEXT
	FOR x = 1 TO MaxX
		FOR y = 1 TO MaxY
			Cell(x, y).Usg = ""
			Cell(x, y).ID = "\"
		NEXT
	NEXT
END SUB

SUB Center (row, text$)
	LOCATE row, 41 - LEN(text$) / 2
	PRINT text$;
END SUB

SUB ChangeDir
		DIM scrn(25, 80), colr(25, 80)

		FOR row = 15 TO 19
			FOR col = 7 TO 59
				scrn(row, col) = SCREEN(row, col)
				colr(row, col) = SCREEN(row, col, 1)
			NEXT
		NEXT

	ShowStat "Cur_DIR"
		COLOR MenuRegular, 0
		Box 15, 7, 19, 59, DblBox
			LOCATE 16, 8: PRINT "Enter name of directory:"
				Kbd$ = GetString$(18, 8, CurrDrive + ":" + RTRIM$(CurrDir), CD$, 51, 128)
					IF MID$(CD$, 2, 1) = ":" THEN
						CurrDrv$ = UCASE$(LEFT$(CD$, 1))
						CurrDr$ = UCASE$(MID$(CD$, 3))
							CALL SubExist(CurrDrv$ + ":" + CurrDr$ + CHR$(0), valid%)
								IF valid% THEN
									CurrDrive = CurrDrv$
									CurrDir = CurrDr$
									GetFreeSpace
								ELSE
									SOUND 3800, 2
									SOUND 32767, 1
								END IF
					END IF
		FOR row = 15 TO 19
			FOR col = 7 TO 59
				LOCATE row, col
					COLOR colr(row, col), 0
					PRINT CHR$(scrn(row, col));
			NEXT
		NEXT
END SUB

SUB CheckInFormula
	FOR col = 1 TO MaxX
		FOR row = 1 TO MaxY
			IF RTRIM$(Cell(col, row).Formula) <> "" THEN CalcForm col, row
		NEXT
	NEXT
END SUB

SUB CheckKeys
	COLOR InfoColor, 0
	CALL GetKbd(ins%, caps%, Num%, scr%)
		LOCATE 25, 64
		IF Num% THEN
			PRINT "NL";
		ELSE
			PRINT "  ";
		END IF
		IF caps% THEN
			PRINT "CL";
		ELSE
			PRINT "  ";
		END IF
COLOR 15, 0
END SUB

SUB ChooseColor (colr)
		DIM scrn(18, 58) AS INTEGER

	FOR row = 16 TO 18
		FOR col = 41 TO 58
			scrn(row, col) = SCREEN(row, col)
		NEXT
	NEXT

	COLOR MenuRegular, 0
	Box 16, 41, 18, 58, DblBox
	LOCATE 17, 42
		FOR col = 0 TO 15
			COLOR 0, col: PRINT "*";
		NEXT

	COLOR 0, colr

	DO
		COLOR 0, colr
		SpinThing 17, colr + 42, 1

		Kbd$ = ""
		DO WHILE Kbd$ = ""
			Kbd$ = INKEY$
				COLOR 0, colr
				SpinThing 17, colr + 42, 1
		LOOP

		SELECT CASE Kbd$
			CASE CHR$(0) + "M"
				colr = colr + 1
				IF colr = 16 THEN colr = 0
				LOCATE 17, 42
					FOR col = 0 TO 15
						COLOR 0, col: PRINT "*";
					NEXT
			CASE CHR$(0) + "K"
				colr = colr - 1
				IF colr = -1 THEN colr = 15
				LOCATE 17, 42
					FOR col = 0 TO 15
						COLOR 0, col: PRINT "*";
					NEXT
			CASE CHR$(13)
				EXIT DO
			CASE ELSE: SOUND 3800, 2: SOUND 32767, 1
		END SELECT
	LOOP

COLOR 15, 0
	FOR row = 16 TO 18
		FOR col = 41 TO 58
			LOCATE row, col
				PRINT CHR$(scrn(row, col));
		NEXT
	NEXT
END SUB

SUB ClearBuffer
	DO UNTIL INKEY$ = "": LOOP
END SUB

SUB ClearSheet
	LOCATE , , 0
	FOR row = 4 TO 23
		CALL XQPrintD(SPACE$(76), row, 4, wattr%, 0)
	NEXT
END SUB

SUB ClearSheetAfterMenu
COLOR 15, 0
LOCATE 2, 1: PRINT SPACE$(80);
LOCATE 3, 1
COLOR 0, SheetColor
PRINT "SP*";
ShowData FirstY, CHR$(FirstX)
IF ShowCell = TRUE THEN ShowCurCell
END SUB

SUB ClearTop
	COLOR 7, 0
	LOCATE 1, 1
	PRINT SPACE$(160);
END SUB

SUB CW
	ShowData FirstY, CHR$(FirstX)
	ShowCurCell
	ClearTop
	COLOR 15, 0
	LOCATE 1, 1
	PRINT "Change the Width of column " + CurrCell + " (" + STR$(W(lt)) + " )"
	lt = (ASC(CurrCell)) - 64
	FirstX = lt + 64
	LOCATE 2, 1
	PRINT "*";
		nw = W(lt)
		nw$ = STR$(nw)
		nw$ = RTRIM$(nw$)
		nw$ = LTRIM$(nw$)
			LOCATE 2, 2: PRINT nw$
		GOSUB ShowNewData
		DO
			Kbd$ = ""
			DO WHILE Kbd$ = ""
				Kbd$ = INKEY$
				ShowExtra
			LOOP

			SELECT CASE Kbd$

				CASE CHR$(0) + "K"
					nw = nw - 1
					IF nw < 9 THEN nw = 9: SOUND 3800, 2: SOUND 32767, 1
					W(lt) = nw
						GOSUB ShowNewData

				CASE CHR$(0) + "M"
					nw = nw + 1
					IF nw > 20 THEN nw = 20: SOUND 3800, 2: SOUND 32767, 1
					W(lt) = nw
						GOSUB ShowNewData

				CASE CHR$(13)
					ShowData FirstY, CHR$(FirstX)
					ShowCurCell
					EXIT DO

				CASE " " TO "~"
					Kbd$ = GetNumeric$(2, 2, Kbd$, nw$, 2, 2)
						IF nw$ <> "" THEN
							nw = VAL(nw$)
								IF nw < 9 THEN nw = 9: SOUND 3800, 2: SOUND 32767, 1
								IF nw > 20 THEN nw = 20: SOUND 3800, 2: SOUND 32767, 1
							W(lt) = nw
							GOSUB ShowNewData
							ShowData FirstY, CHR$(FirstX)
							ShowCurCell
							EXIT DO
						END IF
			END SELECT
LOOP
ClearBuffer
EXIT SUB

ShowNewData:
		FOR y = 1 TO MaxY
			IF LEFT$(Cell(lt, y).ID, 1) = "\" THEN
				Cell(lt, y).Usg = ""
			ELSE
				FOR q = 1 TO MaxY
					Cell(lt, q).CellValue = LTRIM$(Cell(lt, q).CellValue)
					IF LEFT$(Cell(lt, q).ID, 1) = "\" THEN
						Cell(lt, q).Usg = ""
					END IF
				NEXT
			END IF
		NEXT
		ShowData FirstY, CHR$(FirstX)
		ShowCurCell
		ClearTop
			lt = (ASC(CurrCell) - 64)
		COLOR SheetColor, 0: LOCATE 3, LetterMid - 2: PRINT CHR$(27); CurrCell; CHR$(26);
		LOCATE 1, 1
		PRINT "Change the Width of column " + CurrCell + " (" + STR$(W(lt)) + " )"
		LOCATE 2, 1
		PRINT "*";
		nw$ = STR$(nw)
		nw$ = RTRIM$(nw$)
		nw$ = LTRIM$(nw$)
			LOCATE 2, 2: PRINT nw$
RETURN
END SUB

SUB Delete
	DIM choice$(2), ItemRow(2), ItemCol(2)

	SubMenu = TRUE

	choice$(1) = "Delete Row       CTRL-F8 "
	choice$(2) = "Delete Column   SHIFT-F8 "

	Hot$ = "RC"

	ItemRow(1) = 8: ItemCol(1) = 35
	ItemRow(2) = 9: ItemCol(2) = 35

	curr = MMenu(1, 2, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 1: DeleteRow
		CASE 2: DeleteCol
	END SELECT

	SubMenu = FALSE
	ShowData FirstY, CHR$(FirstX)
END SUB

SUB DeleteCol
	ShowStat "WAIT"
	qt = ASC(CurrCell) - 64
	y = 1
	FOR y = 1 TO MaxY
		FOR lt = qt TO MaxX
			Cell(lt, y).CellValue = Cell(lt + 1, y).CellValue
			IF LEFT$(Cell(lt + 1, y).ID, 1) <> "\" THEN
				Cell(lt, y).Usg = Cell(lt + 1, y).Usg
				Cell(lt, y).ID = Cell(lt + 1, y).ID
			ELSEIF LEFT$(Cell(lt + 1, y).ID, 1) = "\" THEN
				Cell(lt, y).Usg = ""
				Cell(lt, y).ID = "\"
			END IF
			Cell(lt, y).Formula = Cell(lt + 1, y).Formula
		NEXT
	NEXT

	ShowData FirstY, CHR$(FirstX)
END SUB

SUB DeleteRow
	ShowStat "WAIT"
	lt = 1
	FOR lt = 1 TO MaxX
		FOR y = currRow TO MaxY
			Cell(lt, y).CellValue = Cell(lt, y + 1).CellValue
			Cell(lt, y).Usg = Cell(lt, y + 1).Usg
			Cell(lt, y).ID = Cell(lt, y + 1).ID
			Cell(lt, y).Formula = Cell(lt, y + 1).Formula
		NEXT
	NEXT
	ShowData FirstY, CHR$(FirstX)
END SUB

SUB DOS
	COLOR 15, 0
	CLS
	PRINT : PRINT : PRINT : PRINT : PRINT : PRINT : PRINT : PRINT "Type EXIT to return to SPREAD-IT!"
	SHELL
	ShowStat "WAIT"
	SLEEP 1
	ClearTop
	ShowStat "WAIT"
	SLEEP 1
	DrawForm
	ShowData FirstY, CHR$(FirstX)
	ShowCurCell
	ShowExtra
	LOCATE 1, 1
	SHELL CurrDrive + ":"
	SHELL "cd " + RTRIM$(CurrDir)
	GetFreeSpace
END SUB

SUB DrawForm
	CLS
	COLOR 0, SheetColor
	LOCATE 3, 1: PRINT SPACE$(80);
	LOCATE 3, 1: PRINT "SP*";
	LOCATE 3, 80: PRINT "*";
	LOCATE 4, 1
	FOR t = 1 TO 20
		PRINT SPACE$(3);
		COLOR 15, 0
			PRINT SPACE$(76);
		COLOR 0, SheetColor
			PRINT " ";
	NEXT
		PRINT SPACE$(80);
		DrawInfo
END SUB

SUB DrawInfo
	COLOR InfoColor, 0
LOCATE 25, 1
PRINT "Cell: <*****>                                                       "; CHR$(179); " SPREAD-IT!";
	ShowCurCell
END SUB

SUB EraseCursor
	LOCATE CellRow, CellCol
	COLOR 15, 0
		PRINT CellData$
END SUB

SUB FileGet (file$)
DIM file$(100), line$(100), row(10), col(10), scrn(22, 62) AS INTEGER, colr(22, 62) AS INTEGER

	CALL FindFirstF(CurrDrive + ":" + RTRIM$(CurrDir) + "\*" + file$ + CHR$(0), attr%, exists%)
		IF exists% THEN
			file$ = "NO FILES FOUND"
			EXIT SUB
		END IF

	SHELL "dir " + CurrDrive + ":" + RTRIM$(CurrDir) + "\*" + file$ + "/b > LIST.TMP"

	OPEN "LIST.TMP" FOR INPUT AS #1
			t = 0
			DO UNTIL EOF(1)
				t = t + 1
				INPUT #1, line$(t)
			LOOP
	CLOSE #1

	KILL "LIST.TMP"

	q = 0
	st = 1

	FOR p = st TO t
		q = q + 1
		file$(q) = line$(p)
	NEXT

FOR t = 1 TO 10
	row(t) = t + 10
NEXT

FOR col = 1 TO 10
	col(col) = 34
NEXT

		FOR row = 4 TO 22
			FOR col = 18 TO 62
				scrn(row, col) = SCREEN(row, col)
				colr(row, col) = SCREEN(row, col, 1)
			NEXT
		NEXT

COLOR MenuRegular, 0
FirstPlace = 1
currItem = 1
	Box 4, 18, 22, 62, DblBox
	Box 10, 32, 21, 47, SngBox
		LOCATE 5, 22: PRINT "File name: [            ]"
		LOCATE 7, 22: PRINT "Directory of " + LEFT$(CurrDrive + ":" + RTRIM$(CurrDir), 26);
		LOCATE 9, 37: PRINT "Files"
GOSUB ShowItems
DO
	Kbd$ = ""
	DO UNTIL Kbd$ <> ""
		Kbd$ = INKEY$
		ShowExtra
	LOOP

	SELECT CASE Kbd$
		CASE CHR$(0) + "P"
			currItem = currItem + 1
				IF currItem > q THEN currItem = q

				IF currItem = 11 THEN
					currItem = 10
					FirstPlace = FirstPlace + 1
						IF FirstPlace + 10 > q THEN FirstPlace = q - 10
				END IF

				GOSUB ShowItems

		CASE CHR$(0) + "H"
			currItem = currItem - 1
				IF currItem = 0 THEN
					currItem = 1
					FirstPlace = FirstPlace - 1
				END IF
				IF FirstPlace < 1 THEN
					FirstPlace = 1
					currItem = 1
				END IF
				GOSUB ShowItems

		CASE " " TO "-"
				LOCATE 4, 34: PRINT SPACE$(12);
			Kbd$ = GetString$(4, 34, Kbd$, done$, 8, 8)
				IF done$ <> "" THEN
					file$ = done$ + file$
					KILL "LIST.TMP"
					EXIT DO
				END IF

		CASE "/" TO "~"
				LOCATE 4, 34: PRINT SPACE$(12);
			Kbd$ = GetString$(4, 34, Kbd$, done$, 8, 8)
				IF done$ <> "" THEN
					file$ = done$ + file$
					KILL "LIST.TMP"
					EXIT DO
				END IF

		CASE CHR$(13)
			file$ = file$(currItem + FirstPlace - 1)
			EXIT DO

		CASE CHR$(27)
			file$ = "ESC"
			EXIT DO
	END SELECT

LOOP
COLOR MenuRegular, 0
		FOR row = 4 TO 22
			FOR col = 18 TO 62
				LOCATE row, col
					COLOR colr(row, col), 0
					PRINT CHR$(scrn(row, col));
			NEXT
		NEXT
EXIT SUB

ShowItems:
	COLOR MenuRegular, 0
	a = 1
	FOR t = FirstPlace TO FirstPlace + 9
		LOCATE row(a), col(a)
		PRINT LEFT$(file$(t) + "     ", 12);
		a = a + 1
	NEXT

	IF currItem > q THEN RETURN
		LOCATE row(currItem), col(currItem)
			COLOR 0, MenuHighlight
			PRINT LEFT$(file$(currItem + FirstPlace - 1) + "     ", 12);
			COLOR MenuRegular, 0
LOCATE 5, 34: PRINT LEFT$(file$(currItem + FirstPlace - 1) + "           ", 12);
RETURN

END SUB

SUB FindRange (In$, Out$)
	LParen = 0
	RParen = 0
		FOR t = 1 TO LEN(In$)
			IF MID$(In$, t, 1) = "[" THEN
				LParen = t + 1
			END IF
			IF MID$(In$, t, 1) = "]" AND LParen <> 0 THEN
				RParen = t - 1
			END IF
		NEXT
	IF LParen = 0 THEN Out$ = In$: EXIT SUB

	lft$ = LEFT$(In$, LParen - 2)
	rgt$ = MID$(In$, RParen + 2, LEN(In$))

	Frm$ = MID$(In$, LParen, RParen - LParen + 1)

Dots = 0
	FOR t = 1 TO LEN(Frm$)
		IF MID$(Frm$, t, 2) = ".." THEN
				Dots = 1
			Oper$ = MID$(Frm$, t - 1, 1)
			Letter1$ = LEFT$(Frm$, 1)
				Letter1 = ASC(Letter1$)
			Letter$ = RIGHT$(Frm$, 3)
				Letter$ = LTRIM$(Letter$)
			Letter2$ = MID$(Frm$, t + 2, 1)
				IF Letter2$ = " " OR Letter2$ = "." THEN
					Letter2$ = MID$(Frm$, t + 3, 1)
				END IF
				Letter2 = ASC(Letter2$)

			Number1 = VAL(MID$(Frm$, 2, 2))
			Number2$ = RIGHT$(Frm$, 2)
				IF LEFT$(Number2$, 1) = Letter2$ THEN
					Number2$ = RIGHT$(Frm$, 1)
				END IF
			Number2 = VAL(Number2$)

			EXIT FOR
		END IF
	NEXT

IF Dots = 0 THEN Out$ = In$: EXIT SUB

	IF Letter1 <> Letter2 AND Number1 <> Number2 THEN Out$ = In$: EXIT SUB

IF Letter1 = Letter2 THEN
	st = 1
	IF Number2 < Number1 THEN st = -1
	FOR t = Number1 TO Number2 STEP st
		Num$ = STR$(t)
		Num$ = RTRIM$(Num$)
		Num$ = LTRIM$(Num$)
			dt$ = dt$ + Letter1$ + Num$
		IF t <> Number2 THEN
			dt$ = dt$ + Oper$
		END IF
		NEXT
END IF

IF Number1 = Number2 THEN
	Num$ = STR$(Number1)
	Num$ = RTRIM$(Num$)
	Num$ = LTRIM$(Num$)
		st = 1
		IF Letter2 < Letter1 THEN st = -1
		FOR t = Letter1 TO Letter2 STEP st
			Letter$ = CHR$(t)
				dt$ = dt$ + Letter$ + Num$
			IF t <> Letter2 THEN
				dt$ = dt$ + Oper$
			END IF
		NEXT
END IF

Out$ = lft$ + dt$ + rgt$
END SUB

SUB GetCurrStuff
		SHELL "cd > SP!$~TMP.TMP"
			OPEN "SP!$~TMP.TMP" FOR INPUT AS #1
				INPUT #1, stuff$
			CLOSE #1
			CurrDrive = LEFT$(stuff$, 1)
			CurrDir = MID$(stuff$, 3, LEN(stuff$))
		KILL "SP!$~TMP.TMP"
	GetFreeSpace
END SUB

SUB GetEnd
			k$ = ""
			DO WHILE k$ = ""
				k$ = INKEY$
				ShowExtra
				LOCATE CellRow, LetterMid - 1, 1
				FOR cursorCount = 1 TO 5000
					a$ = INKEY$
					IF a$ <> "" THEN k$ = a$: EXIT FOR
				NEXT
				LOCATE , , 0
			LOOP

			SELECT CASE k$
				CASE CHR$(0) + "M"
					FirstX = MaxX + 64
					CurrCell = CHR$(MaxX + 64)

				CASE CHR$(0) + "K"
					FirstX = 65
					CurrCell = "A"

				CASE CHR$(0) + "H"
					FirstY = 1
					currRow = 1

				CASE CHR$(0) + "P"
					FirstY = MaxY
					currRow = MaxY

				CASE CHR$(0) + "O"
					FirstY = MaxY
					currRow = MaxY
					CurrCell = CHR$(MaxX + 64)
					FirstX = MaxX + 64

				CASE ELSE
			END SELECT
END SUB

SUB GetEndOfLine (col, row, Value$)
IF Cell(col, row).ID = "\" THEN EXIT SUB

	Value$ = ""
	FOR t = 1 TO FirstX - 65
		TotWidth = TotWidth + W(t)
	NEXT
	TotWidth = TotWidth + 1

	'IF RTRIM$(Cell(col, row).CellValue) <> "" THEN
		Value$ = MID$(RTRIM$(Cell(col, row).CellValue), TotWidth, 80)
	'END IF
END SUB

SUB GetFormula
	Changed = 1
	CurrNum$ = STR$(currRow)
	CurrNum$ = LTRIM$(CurrNum$)
	CurrNum$ = RTRIM$(CurrNum$)
	LOCATE 2, 1: PRINT CurrCell; CurrNum$; ":";
		lt = (ASC(CurrCell)) - 64
		dt = currRow
		place = 4
		IF LEN(CurrNum$) = 2 THEN place = 5
		IF LEN(CurrNum$) = 3 THEN place = 6
		Kbd$ = GetString$(2, place, RTRIM$(Cell(lt, dt).Formula), Cell(lt, dt).Formula, 80 - place, 100)
			Cell(lt, dt).Formula = UCASE$(Cell(lt, dt).Formula)
			Cell(lt, dt).Formula = LTRIM$(Cell(lt, dt).Formula)
				IF RTRIM$(Cell(lt, dt).Formula) <> "" THEN
					Cell(lt, dt).CellValue = ""
				ELSEIF Cell(lt, dt).Formula = "" THEN
					Cell(lt, dt).Usg = ""
					Cell(lt, dt).ID = "\"
					Cell(lt, dt).CellValue = ""
					Cell(lt, dt).Formula = ""
						ClearTop
						ShowStat "READY"
						IF CalcMode <> "Manual    " THEN
							CheckInFormula
						ELSE
							CalcForm lt, dt
							ShowData FirstY, CHR$(FirstX)
							ShowCurCell
						END IF
						ShowData FirstY, CHR$(FirstX)
						ShowCurCell
					RETURN
				END IF
	ClearTop
	ShowStat "READY"
	IF CalcMode <> "Manual    " THEN CheckInFormula
	Clean = 1
	ShowData FirstY, CHR$(FirstX)
	ShowCurCell
END SUB

SUB GetFreeSpace
DIM line$(1000)

		SHELL "echo. > " + CurrDrive + ":\JUNK.TMP"
		SHELL "dir " + CurrDrive + ": > SP!$~TMP.TMP"

		OPEN "SP!$~TMP.TMP" FOR INPUT AS #1
			t = 0
			DO UNTIL EOF(1)
				t = t + 1
				INPUT #1, line$(t)
			LOOP
		CLOSE #1

		line$(t) = LTRIM$(line$(t))

		line$(t) = LEFT$(line$(t), LEN(line$(t)) - 11)

		FreeSpace = line$(t)

		KILL "SP!$~TMP.TMP"
		KILL CurrDrive + ":\JUNK.TMP"
END SUB

SUB GetNumber
	Changed = 1
	CurrNum$ = STR$(currRow)
	CurrNum$ = LTRIM$(CurrNum$)
	CurrNum$ = RTRIM$(CurrNum$)
		LOCATE 2, 1: PRINT CurrCell; CurrNum$; ":";
		place = 4
		IF LEN(CurrNum$) = 2 THEN place = 5
		IF LEN(CurrNum$) = 3 THEN place = 6
		lt = ASC(CurrCell) - 64
		dt = currRow
				Kbd$ = GetNumeric$(2, place, Kbd$, Cell(lt, dt).CellValue, 80 - place, 100)
					IF LEN(RTRIM$(Cell(lt, dt).CellValue)) > 8 AND LEFT$(Cell(lt, dt).ID, 1) <> "C" AND LEFT$(NumType, 1) <> "C" THEN
						rt$ = NumType
						NumType = "S" + Decimal
						Translate RTRIM$(Cell(lt, dt).CellValue), NewUs$
						Cell(lt, dt).Usg = NewUs$
						Cell(lt, dt).ID = NumType
						NumType = rt$
					ELSEIF RTRIM$(Cell(lt, dt).CellValue) <> "" THEN
						Translate (Cell(lt, dt).CellValue), NewUs$
						Cell(lt, dt).Usg = NewUs$
						Cell(lt, dt).ID = NumType
					ELSEIF Cell(lt, dt).CellValue = "" THEN
						Cell(lt, dt).Usg = ""
						Cell(lt, dt).ID = "\"
						Cell(lt, dt).CellValue = ""
						Cell(lt, dt).Formula = ""
							ClearTop
							ShowStat "READY"
							IF CalcMode <> "Manual    " THEN CheckInFormula
							ShowData FirstY, CHR$(FirstX)
							ShowCurCell
						RETURN
					END IF
	ClearTop
	ShowStat "READY"
	IF CalcMode <> "Manual    " THEN CheckInFormula
	Clean = 1
	ShowData FirstY, CHR$(FirstX)
	ShowCurCell
END SUB

FUNCTION GetNumeric$ (row, col, Start$, end$, Vis, Max)
	ShowStat "EDIT"
	curr$ = LTRIM$(LEFT$(Start$, Max))
	IF curr$ = CHR$(8) THEN curr$ = ""

	LOCATE , , 1

	finished = FALSE
	DO
		GOSUB GetNumShowText
		GOSUB GetNumGetKey

			IF LEN(Kbd$) > 1 THEN
				Move$ = Kbd$
					finished = TRUE
					GetNumeric$ = Kbd$
					LOCATE , , 0
		ELSE
			SELECT CASE Kbd$
				CASE CHR$(13), CHR$(27)

					finished = TRUE
					GetNumeric$ = Kbd$
					LOCATE , , 0

				CASE CHR$(8)
					IF curr$ <> "" THEN
						curr$ = LEFT$(curr$, LEN(curr$) - 1)
					END IF
				CASE "0" TO "9"
					IF LEN(curr$) < Max THEN
						curr$ = curr$ + Kbd$
					END IF
				CASE "-"
					IF LEN(curr$) = 0 THEN
						curr$ = curr$ + Kbd$
					END IF
				CASE "."
					IF LEN(curr$) < Max THEN
						curr$ = curr$ + Kbd$
					END IF
				CASE ELSE
					SOUND 3800, 2
					SOUND 32767, 1
			END SELECT
		END IF

	LOOP UNTIL finished

	end$ = curr$

	LOCATE , , 0
	EXIT FUNCTION

GetNumShowText:
	LOCATE row, col
	IF LEN(curr$) > Vis THEN
		PRINT RIGHT$(curr$, Vis);
	ELSE
		PRINT curr$; SPACE$(Vis - LEN(curr$));
		LOCATE row, col + LEN(curr$)
	END IF
	RETURN

GetNumGetKey:
	Kbd$ = ""
	WHILE Kbd$ = ""
			ShowExtra
		Kbd$ = INKEY$
			LOCATE row, (col + LEN(curr$)), 1
		FOR t = 1 TO 20000: NEXT
	WEND
	RETURN
END FUNCTION

FUNCTION GetString$ (row, col, Start$, end$, Vis, Max)
	ShowStat "EDIT"
	curr$ = LEFT$(Start$, Max)
	IF curr$ = CHR$(8) THEN curr$ = ""

	LOCATE , , 1

	place = 2 + LEN(Start$)
	PPlace = LEN(Start$)
		IF PPlace = 0 THEN PPlace = 1

	finished = FALSE
	DO
		GOSUB GetStringShowText
		GOSUB GetStringGetKey

			IF LEN(Kbd$) > 1 THEN
				Move$ = Kbd$
					finished = TRUE
					GetString$ = Kbd$
					LOCATE , , 0
			ELSE

			SELECT CASE Kbd$
				CASE CHR$(13), CHR$(27)
					finished = TRUE
					GetString$ = Kbd$
					LOCATE , , 0

				CASE CHR$(8)
					IF curr$ <> "" THEN
						curr$ = LEFT$(curr$, LEN(curr$) - 1)
					END IF

				CASE " " TO "~"
					IF LEN(curr$) < Max THEN
						curr$ = curr$ + Kbd$
					END IF

				CASE ELSE
					SOUND 3800, 2
					SOUND 32767, 1
			END SELECT
		END IF

	LOOP UNTIL finished

	end$ = curr$

	LOCATE , , 0
	EXIT FUNCTION

GetStringShowText:
	LOCATE row, col
	IF LEN(curr$) > Vis THEN
		PRINT RIGHT$(curr$, Vis);
	ELSE
		PRINT curr$; SPACE$(Vis - LEN(curr$));
		LOCATE row, col + LEN(curr$)
	END IF
	RETURN

GetStringGetKey:
	Kbd$ = ""
	WHILE Kbd$ = ""
			ShowExtra
		Kbd$ = INKEY$
			LOCATE row, (col + LEN(curr$)), 1
		FOR t = 1 TO 20000: NEXT
	WEND
	RETURN
END FUNCTION

SUB GetText
	Changed = 1
	CurrNum$ = STR$(currRow)
	CurrNum$ = LTRIM$(CurrNum$)
	CurrNum$ = RTRIM$(CurrNum$)
	LOCATE 2, 1: PRINT CurrCell; CurrNum$; ":";
		place = 4
		IF LEN(CurrNum$) = 2 THEN place = 5
		IF LEN(CurrNum$) = 3 THEN place = 6
			Kbd$ = GetString$(2, place, Kbd$, Cell((ASC(CurrCell)) - 64, currRow).CellValue, 80 - place, 100)
				Cell(ASC(CurrCell) - 64, currRow).Usg = ""
					IF Cell(ASC(CurrCell) - 64, currRow).CellValue = "" THEN
						Cell(ASC(CurrCell) - 64, currRow).Formula = ""
						Cell(ASC(CurrCell) - 64, currRow).ID = "\"
							ClearTop
							ShowStat "READY"
							IF CalcMode <> "Manual    " THEN CheckInFormula
							ShowData FirstY, CHR$(FirstX)
							ShowCurCell
						RETURN
					END IF
	ClearTop
	ShowStat "READY"
	IF CalcMode <> "Manual    " THEN CheckInFormula
	Clean = 1
	ShowData FirstY, CHR$(FirstX)
	ShowCurCell
END SUB

SUB Insert
	DIM choice$(2), ItemRow(2), ItemCol(2)
	SubMenu = TRUE

	choice$(1) = "Insert Row          F8 "
	choice$(2) = "Insert Column   ALT-F8 "

	Hot$ = "RC"

	ItemRow(1) = 8: ItemCol(1) = 35
	ItemRow(2) = 9: ItemCol(2) = 35

	curr = MMenu(1, 2, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 1: InsertRow
		CASE 2: InsertCol
	END SELECT

	SubMenu = FALSE
	ShowData FirstY, CHR$(FirstX)
END SUB

SUB InsertCol
	ShowStat "WAIT"
	qt = ASC(CurrCell) - 64
	y = 1
	FOR y = 1 TO MaxY
		FOR lt = MaxX TO qt STEP -1
			Cell(lt, y).CellValue = Cell(lt - 1, y).CellValue
			Cell(lt, y).Usg = Cell(lt - 1, y).Usg
			Cell(lt, y).Formula = Cell(lt - 1, y).Formula
			IF LEFT$(Cell(lt - 1, y).ID, 1) <> "\" THEN
				Cell(lt, y).Usg = Cell(lt - 1, y).Usg
				Cell(lt, y).ID = Cell(lt - 1, y).ID
			ELSEIF LEFT$(Cell(lt - 1, y).ID, 1) = "\" THEN
				Cell(lt, y).Usg = ""
				Cell(lt, y).ID = "\"
			END IF
		NEXT
	NEXT

	FOR lt = 1 TO MaxY
		Cell(qt, lt).CellValue = ""
		Cell(qt, lt).Usg = ""
		Cell(qt, lt).ID = "\"
		Cell(qt, lt).Formula = ""
	NEXT

	ShowData FirstY, CHR$(FirstX)
	Changed = 1
END SUB

SUB InsertRow
	ShowStat "WAIT"
	FOR lt = 1 TO MaxX
		FOR y = MaxY TO currRow STEP -1
			Cell(lt, y).CellValue = Cell(lt, y - 1).CellValue
			Cell(lt, y).Usg = Cell(lt, y - 1).Usg
			Cell(lt, y).ID = Cell(lt, y - 1).ID
			Cell(lt, y).Formula = Cell(lt, y - 1).Formula
		NEXT
	NEXT

	FOR lt = 1 TO MaxX
		Cell(lt, currRow).CellValue = ""
		Cell(lt, currRow).Usg = ""
		Cell(lt, currRow).ID = "\"
		Cell(lt, currRow).Formula = ""
	NEXT

	ShowData FirstY, CHR$(FirstX)
	Changed = 1
END SUB

SUB Intro
	LoadINI
	PRINT : PRINT "SPREAD-IT! Version 3.00"
			PRINT "(c)Copyright 1993 by Nathan Thomas"
	SLEEP 3
	GetCurrStuff
		DefaultDir$ = CurrDrive + ":" + RTRIM$(CurrDir)
	DrawForm
		ShowCurCell
		ShowExtra
		ClearBuffer
		CalcHalves
END SUB

SUB LoadINI
	OPEN "SPREADIT.INI" FOR INPUT AS #1
		INPUT #1, MenuHighlight
		INPUT #1, MenuRegular
		INPUT #1, MenuHotKey
		INPUT #1, SheetColor
		INPUT #1, StatColor
		INPUT #1, InfoColor
		INPUT #1, CursorColor
		INPUT #1, DefaultFile
		INPUT #1, CalcMode
		INPUT #1, Order
		INPUT #1, Iter
	CLOSE #1
END SUB

SUB MainSystem
	DIM choice$(11), ItemRow(11), ItemCol(11), scrn(25, 80)  AS INTEGER, colr(25, 80) AS INTEGER
		status = 0
		NewSpread
		IF DefaultFile <> "" THEN
			CurrSheet = DefaultFile
				CALL exist(DefaultFile + ".SHT" + CHR$(0), exists%)
					IF NOT exists% THEN status = 1
			IF status = 0 THEN
				OPEN CurrDrive + ":" + RTRIM$(CurrDir) + "\" + CurrSheet FOR INPUT AS #1
					FOR col = 1 TO MaxX
						INPUT #1, W(col)
						FOR row = 1 TO MaxY
							LINE INPUT #1, Cell(col, row).Usg
							LINE INPUT #1, Cell(col, row).Formula
							LINE INPUT #1, Cell(col, row).CellValue
								 INPUT #1, Cell(col, row).ID
						NEXT
					NEXT
					INPUT #1, NumType
			END IF
		END IF
	ShowExtra
	ShowData 1, "A"
	ShowCurCell
		LOCATE , , 0
		ShowStat "READY"
		GOSUB ManipulateSheet

MenuSystem:
MainCurr = 1
DO
		choice$(1) = " File "
		choice$(2) = " Edit "
		choice$(3) = " Style "
		choice$(4) = " Options "
		choice$(5) = " Tools "

		Hot$ = "FESOT"

		ItemRow(1) = 1: ItemCol(1) = 1
		ItemRow(2) = 1: ItemCol(2) = 7
		ItemRow(3) = 1: ItemCol(3) = 13
		ItemRow(4) = 1: ItemCol(4) = 20
		ItemRow(5) = 1: ItemCol(5) = 29

		ClearTop
		EraseCursor
		MC = 0
		MainCurr = MMenu(MainCurr, 5, choice$(), ItemRow(), ItemCol(), Hot$, TRUE)

FindMain:
			SELECT CASE MainCurr
				CASE 0: MainCurr = 1: EXIT DO
				CASE 1: GOSUB FileSystem
				CASE 2: GOSUB EditSystem
				CASE 3: GOSUB StyleSystem
				CASE 4: GOSUB OptionSystem
				CASE 5: GOSUB ToolSystem
			END SELECT
	IF MC <> 0 THEN MainCurr = MC: GOTO FindMain
LOOP
ShowCell = TRUE
ClearSheetAfterMenu
ShowCell = FALSE
MainCurr = 0
RETURN

FileSystem:
	MC = 0
	choice$(1) = " New Spreadsheet          F7 "
	choice$(2) = " Open Spreadsheet         F6" + CHR$(16)
	choice$(3) = " Save                 ALT-F6"
		IF CurrSheet = "" THEN
			choice$(3) = choice$(3) + CHR$(16)
		ELSE
			choice$(3) = choice$(3) + " "
		END IF
	choice$(4) = " Save As             CTRL-F6" + CHR$(16)
	choice$(5) = " Close              SHIFT-F6 "
	choice$(6) = " Erase                       "
	choice$(7) = "*****************************"
	choice$(8) = " DOS Shell            ALT-F7 "
	choice$(9) = "*****************************"
	choice$(10) = " Directory            ALT-F5 "
	choice$(11) = " Exit                     F5" + CHR$(16)

	Hot$ = "NOSACE D ix"

	ItemRow(1) = 3: ItemCol(1) = 2
	ItemRow(2) = 4: ItemCol(2) = 2
	ItemRow(3) = 5: ItemCol(3) = 2
	ItemRow(4) = 6: ItemCol(4) = 2
	ItemRow(5) = 7: ItemCol(5) = 2
	ItemRow(6) = 8: ItemCol(6) = 2
	ItemRow(7) = 9: ItemCol(7) = 2
	ItemRow(8) = 10: ItemCol(8) = 2
	ItemRow(9) = 11: ItemCol(9) = 2
	ItemRow(10) = 12: ItemCol(10) = 2
	ItemRow(11) = 13: ItemCol(11) = 2

	curr = MMenu(1, 11, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE -3: ClearSheetAfterMenu: ShowMenu 2: MC = 2: MenuLeft = 7: MenuRight = 12: RETURN
		CASE -2: ClearSheetAfterMenu: ShowMenu 5: MC = 5: MenuLeft = 29: MenuRight = 35: RETURN
		CASE 1: NewSpread: ClearBuffer: ClearSheetAfterMenu: ShowMenu 1: GOTO FileSystem
		CASE 2: OpenSpread: ClearBuffer: GOTO FileSystem
		CASE 3: SaveSpread: ClearBuffer: GOTO FileSystem
		CASE 4: CurrSheet = "": SaveSpread: ClearBuffer: GOTO FileSystem
		CASE 5:
			SaveSpread
			IF file$ <> "ESC" THEN NewSpread
			ClearBuffer
			GOTO FileSystem
		CASE 6: KILL CurrDrive + ":" + RTRIM$(CurrDir) + "\" + CurrSheet: ClearBuffer: GOTO FileSystem
		CASE 8: DOS: ClearBuffer: ShowMenu 1: GOTO FileSystem
		CASE 10: ChangeDir: ClearBuffer: GOTO FileSystem
		CASE 11: GOSUB ExitSystem: ClearBuffer: GOTO FileSystem
	END SELECT

	ClearBuffer
	ClearSheetAfterMenu
	RETURN

ToolSystem:
	MC = 0
	choice$(1) = " Set Defaults   " + CHR$(16)

	Hot$ = "D"

	ItemRow(1) = 3: ItemCol(1) = 30

	curr = MMenu(1, 1, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE -2: ClearSheetAfterMenu: ShowMenu 4: MC = 4: MenuLeft = 20: MenuRight = 28: RETURN
		CASE -3: ClearSheetAfterMenu: ShowMenu 1: MC = 1: MenuLeft = 1: MenuRight = 6: RETURN
		CASE 1: GOSUB SetDefaults: ClearBuffer: ClearSheetAfterMenu: ShowMenu 5: GOTO ToolSystem
	END SELECT

	ClearSheetAfterMenu
	RETURN

ExitSystem:
	IF Changed = 0 THEN GOTO FancyFinish
	SubMenu = TRUE

	choice$(1) = "* Lose your changes and Exit? *"
	choice$(2) = " No                            "
	choice$(3) = " Yes                           "
	choice$(4) = " Save and Exit                 "

	Hot$ = " NYS"

	ItemRow(1) = 8: ItemCol(1) = 13
	ItemRow(2) = 9: ItemCol(2) = 13
	ItemRow(3) = 10: ItemCol(3) = 13
	ItemRow(4) = 11: ItemCol(4) = 13

			FOR row = 7 TO 12
				FOR col = 12 TO 44
					colr(row, col) = SCREEN(row, col, 1)
					scrn(row, col) = SCREEN(row, col)
				NEXT
			NEXT

	curr = MMenu(2, 4, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

			FOR row = 7 TO 12
				FOR col = 12 TO 44
					LOCATE row, col
						COLOR colr(row, col), 0
							PRINT CHR$(scrn(row, col))
				NEXT
			NEXT

	SELECT CASE curr
		CASE 3: GOTO FancyFinish
		CASE 4: SaveSpread: GOTO FancyFinish
	END SELECT

	ClearBuffer
	SubMenu = FALSE
	RETURN

OptionSystem:
	MC = 0
	choice$(1) = " Colors          ALT-F4" + CHR$(16)
	choice$(2) = " Recalculation         " + CHR$(16)

	Hot$ = "CR"

	ItemRow(1) = 3: ItemCol(1) = 21
	ItemRow(2) = 4: ItemCol(2) = 21

	curr = MMenu(1, 2, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE -2: ClearSheetAfterMenu: ShowMenu 3: MC = 3: MenuLeft = 13: MenuRight = 19: RETURN
		CASE -3: ClearSheetAfterMenu: ShowMenu 5: MC = 5: MenuLeft = 29: MenuRight = 35: RETURN
		CASE 1: OptionMenu = TRUE: GOSUB ChangeColors: OptionMenu = FALSE: ClearBuffer: ShowData FirstY, CHR$(FirstX): GOTO OptionSystem
		CASE 2: ReCalc: ClearBuffer: GOTO OptionSystem
	END SELECT

	ClearBuffer
	ClearSheetAfterMenu
	RETURN

FancyFinish:
	ClearSheetAfterMenu
	ClearSheet
		SLEEP 1
		LOCATE 24, 1: PRINT SPACE$(79);
		LOCATE 25, 1: PRINT SPACE$(80);
		SLEEP 1
		col = 80
			FOR row = 3 TO 24
				LOCATE row, col: PRINT " ";
			NEXT
		SLEEP 1
		ClearTop
	SLEEP 1
	LOCATE 3, 1: PRINT SPACE$(80);
		col = 1
			FOR row = 3 TO 23
				LOCATE row, col: PRINT "   ";
			NEXT
		SLEEP 1
	SaveINI
	LOCATE 1, 1
	SLEEP 1
	EXIT SUB

SetDefaults:
	SubMenu = TRUE
	ShowMenu 5

	choice$(1) = " Startup File   " + RIGHT$(SPACE$(13) + RTRIM$(DefaultFile), 12) + " " + CHR$(16)
	choice$(2) = " Reset Colors          ALT-F4 "
	choice$(3) = " Reset Widths          ALT-F3 "

	Hot$ = "SCW"

	ItemRow(1) = 7: ItemCol(1) = 30
	ItemRow(2) = 8: ItemCol(2) = 30
	ItemRow(3) = 9: ItemCol(3) = 30

	curr = MMenu(1, 3, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 1: SetFile: ClearSheetAfterMenu: ShowMenu 5: ClearBuffer: GOTO SetDefaults
		CASE 2: GOSUB SetColors: ClearBuffer: GOTO SetDefaults
		CASE 3: SetWidths: ShowData FirstY, CHR$(FirstX): ShowMenu 5: ClearBuffer: GOTO SetDefaults
	END SELECT

	ClearBuffer
	SubMenu = FALSE
	RETURN

ChangeColors:
	SubMenu = TRUE

	choice$(1) = " Menu Colors       " + CHR$(16)
	choice$(2) = " Sheet Color        "
	choice$(3) = " Status Color       "
	choice$(4) = " Information Color  "
	choice$(5) = " Cursor Color       "

	Hot$ = "MStICH"

	ItemRow(1) = 10: ItemCol(1) = 50
	ItemRow(2) = 11: ItemCol(2) = 50
	ItemRow(3) = 12: ItemCol(3) = 50
	ItemRow(4) = 13: ItemCol(4) = 50
	ItemRow(5) = 14: ItemCol(5) = 50

	curr = MMenu(1, 5, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 1: GOSUB SetMenuColors: ClearBuffer: GOTO ChangeColors
		CASE 2: GOSUB SetSheet: ClearBuffer: GOTO ChangeColors
		CASE 3: GOSUB SetStat: ClearBuffer: GOTO ChangeColors
		CASE 4: GOSUB SetInfo: ClearBuffer: GOTO ChangeColors
		CASE 5: GOSUB SetCursor: ClearBuffer: GOTO ChangeColors
	END SELECT

	ClearBuffer
	SubMenu = FALSE
	RETURN

SetMenuColors:
	ShowMenu 6
	IF OptionMenu = TRUE THEN ShowMenu 4

	choice$(1) = " Highlight  "
	choice$(2) = " Normal     "
	choice$(3) = " Hot Key    "

	Hot$ = "HNK"

	ItemRow(1) = 13: ItemCol(1) = 45
	ItemRow(2) = 14: ItemCol(2) = 45
	ItemRow(3) = 15: ItemCol(3) = 45

		IF Chng = 1 THEN
			FOR row = 12 TO 16
				FOR col = 44 TO 57
					colr(row, col) = SCREEN(row, col, 1)
					scrn(row, col) = SCREEN(row, col)
				NEXT
			NEXT
			Chng = 0
		END IF

	curr = MMenu(1, 3, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 1: GOSUB SetHigh: ClearBuffer: GOTO SetMenuColors
		CASE 2: GOSUB SetNorm: ClearBuffer: GOTO SetMenuColors
		CASE 3: GOSUB SetHot: ClearBuffer: GOTO SetMenuColors
	END SELECT

	ClearBuffer
			FOR row = 12 TO 16
				FOR col = 44 TO 57
					LOCATE row, col
						COLOR colr(row, col), 0
							PRINT CHR$(scrn(row, col));
				NEXT
			NEXT
	Chng = 1
	RETURN

SetHigh:
	ShowStat "HIGH"
		ChooseColor MenuHighlight

	RETURN

SetNorm:
	ShowStat "NORM"
		ChooseColor MenuRegular
			ShowMenu 6
			IF OptionMenu = TRUE THEN ShowMenu 4
	RETURN

SetHot:
	ShowStat "HOTKEY"
		ChooseColor MenuHotKey
	RETURN

SetSheet:
	ShowStat "SHEET"
		ChooseColor SheetColor
			DrawForm
			ShowData FirstY, CHR$(FirstX)
				ShowMenu 6
				IF OptionMenu = TRUE THEN ShowMenu 4
	RETURN

SetStat:
	ShowStat "STATUS"
		ChooseColor StatColor
	RETURN

SetInfo:
	ShowStat "INFO"
		ChooseColor InfoColor
		DrawInfo
	RETURN

SetCursor:
	ShowStat "CURSOR"
		ChooseColor CursorColor
	RETURN

SetColors:
MenuHighlight = 7
MenuRegular = 7
MenuHotKey = 15
SheetColor = 15
StatColor = 15
InfoColor = 15
CursorColor = 7
RETURN

EditSystem:
	MC = 0
	choice$(1) = " Insert       "
	choice$(2) = " Delete       "

	Hot$ = "ID"

	ItemRow(1) = 3: ItemCol(1) = 8
	ItemRow(2) = 4: ItemCol(2) = 8

	curr = MMenu(1, 2, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE -2: ClearSheetAfterMenu: ShowMenu 1: MC = 1: MenuLeft = 1: MenuRight = 6: RETURN
		CASE -3: ClearSheetAfterMenu: ShowMenu 3: MC = 3: MenuLeft = 13: MenuRight = 19: RETURN
		CASE 1: Insert: ClearBuffer: GOTO EditSystem
		CASE 2: Delete: ClearBuffer: GOTO EditSystem
	END SELECT
	ClearBuffer
	ClearSheetAfterMenu
	RETURN

StyleSystem:
	MC = 0
	choice$(1) = " Numeric Format     F2" + CHR$(16)
	choice$(2) = "***********************"
	choice$(3) = " Column Width       F3" + CHR$(16)
	choice$(4) = " Reset Widths   ALT-F3 "

	Hot$ = "F WR"

	ItemRow(1) = 3: ItemCol(1) = 14
	ItemRow(2) = 4: ItemCol(2) = 14
	ItemRow(3) = 5: ItemCol(3) = 14
	ItemRow(4) = 6: ItemCol(4) = 14

	curr = MMenu(1, 4, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE -2: ClearSheetAfterMenu: ShowMenu 2: MC = 2: MenuLeft = 7: MenuRight = 12: RETURN
		CASE -3: ClearSheetAfterMenu: ShowMenu 4: MC = 4: MenuLeft = 20: MenuRight = 28: RETURN
		CASE 1: SelectFormat: ShowMenu 3: ClearBuffer: GOTO StyleSystem
		CASE 3: CW: EraseCursor: ShowMenu 3: ClearBuffer: GOTO StyleSystem
		CASE 4: SetWidths: ClearBuffer: GOTO StyleSystem
	END SELECT

	ClearBuffer
	ClearSheetAfterMenu
	RETURN

ManipulateSheet:
DO
		Check = 1
		Kbd$ = ""
		ShowStat "READY"
		LOCATE CellRow, LetterMid - 1, 1
		DO WHILE Kbd$ = ""
			Kbd$ = INKEY$
				IF Move$ <> "" THEN
					Kbd$ = Move$
					Move$ = ""
				END IF
			ShowExtra
			FOR cursorCount = 1 TO 20000: NEXT
		LOOP
		LOCATE , , 0
		Check = 0

	SELECT CASE Kbd$

		CASE CHR$(0) + "<"                  'F2
			SelectFormat
			ShowStat "READY"
			ShowCurCell
			Move$ = ""

		CASE CHR$(0) + "="                  'F3
			CW
			ClearTop
			ShowStat "READY"
			ShowCurCell
			Move$ = ""

		CASE CHR$(0) + ">"                  'F4
			lt = (ASC(CurrCell) - 64)
			dt = currRow
			Us$ = LEFT$(Cell(lt, dt).ID, 1)
				Kbd$ = RTRIM$(Cell(lt, dt).CellValue)
				IF Us$ = "S" OR Us$ = "F" OR Us$ = "C" OR Us$ = "P" THEN GetNumber
				IF Us$ = "\" THEN GetText

		CASE CHR$(0) + "?"                  'F5
			GOSUB ExitSystem
			ShowStat "READY"
			ShowCurCell

		CASE CHR$(0) + "j"                  'ALT-F3
			SetWidths
			ShowStat "READY"
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell

		CASE CHR$(0) + "k"                  'ALT-F4
			GOSUB ChangeColors
			ShowStat "READY"
			ShowCurCell

		CASE CHR$(0) + "a"                  'CTRL-F4
			GOSUB SetColors
			ShowStat "READY"

		CASE CHR$(0) + "l"                  'ALT-F5
			ChangeDir
			ShowStat "READY"
			ShowCurCell
			Move$ = ""

		CASE CHR$(0) + "@"                  'F6
			OpenSpread
			ShowStat "READY"
			Move$ = ""

		CASE CHR$(0) + "m"                  'ALT-F6
			SaveSpread
			ShowStat "READY"
			Move$ = ""

		CASE CHR$(0) + "c"                  'CTRL-F6
			CurrSheet = ""
			SaveSpread
			ShowStat "READY"
			ShowCurCell
			Move$ = ""

		CASE CHR$(0) + "Y"                  'SHIFT-F6
			SaveSpread
			NewSpread
			ShowStat "READY"
			ShowCurCell
			Move$ = ""

		CASE CHR$(0) + "A"                  'F7
			NewSpread
			ShowStat "READY"

		CASE CHR$(0) + "n"                  'ALT-F7
			DOS
			ShowStat "READY"
			ShowCurCell

		CASE CHR$(0) + "B"                  'F8
			InsertRow
			ShowStat "READY"

		CASE CHR$(0) + "C"                  'F9
			lt = (ASC(CurrCell) - 64)
			dt = currRow
				CalcForm lt, dt
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell

		CASE CHR$(0) + "o"                  'ALT-F8
			InsertCol
			ShowStat "READY"

		CASE CHR$(0) + "e"                  'CTRL-F8
			DeleteRow
			ShowStat "READY"

		CASE CHR$(0) + "["                  'SHIFT-F8
			DeleteCol
			ShowStat "READY"

		CASE CHR$(0) + "S"                  'DEL
			COLOR InfoColor, 0
			LOCATE 25, 49: PRINT "DEL"; : FOR t = 1 TO 15000: NEXT
			LOCATE 25, 49: PRINT "   ";
				Cell(ASC(CurrCell) - 64, currRow).Usg = ""
				Cell(ASC(CurrCell) - 64, currRow).ID = "\"
				Cell(ASC(CurrCell) - 64, currRow).CellValue = ""
				Cell(ASC(CurrCell) - 64, currRow).Formula = ""
			IF CalcMode <> "Manual    " THEN CheckInFormula
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell
			Changed = 1

		CASE CHR$(0) + "H"                  'UP
			EraseCursor
			currRow = currRow - 1
				IF currRow = 0 THEN
					currRow = 1
					SOUND 3800, 3
				END IF
				IF currRow < FirstY THEN
					FirstY = FirstY - 1
					ShowData FirstY, CHR$(FirstX)
					ShowCurCell
				ELSE
					ShowLetters CHR$(FirstX)
					ShowNumbers FirstY
					ShowCurCell
				END IF


		CASE CHR$(0) + "P"                  'DOWN
			EraseCursor
			currRow = currRow + 1
				IF currRow > MaxY THEN
					currRow = MaxY
					SOUND 3800, 3
				END IF
				IF currRow > FirstY + 19 THEN
					FirstY = FirstY + 1
						ShowData FirstY, CHR$(FirstX)
						ShowCurCell
				ELSE
					ShowLetters CHR$(FirstX)
					ShowNumbers FirstY
					ShowCurCell
				END IF

		CASE CHR$(0) + "Q"                  'Page Down
			currRow = currRow + 20
			FirstY = FirstY + 20
				IF currRow > MaxY THEN currRow = MaxY
				IF FirstY > MaxY THEN FirstY = MaxY
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell

		CASE CHR$(0) + "I"                  'Page Up
			currRow = currRow - 20
				FirstY = FirstY - 20
				IF currRow < 1 THEN currRow = 1
				IF FirstY < 1 THEN FirstY = 1
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell

		CASE CHR$(0) + "M"                  'RIGHT
			currC = (ASC(CurrCell)) - 64
				EraseCursor
				currC = currC + 1
					IF currC > MaxX THEN
						currC = MaxX
						SOUND 3800, 3
					END IF
					IF currC > (ASC(LastChar) - 64) THEN
						FirstX = FirstX + 1
							CurrCell = (CHR$(currC + 64))
							ShowData FirstY, CHR$(FirstX)
							ShowCurCell
					ELSE
						CurrCell = (CHR$(currC + 64))
						ShowLetters CHR$(FirstX)
						ShowNumbers FirstY
						ShowCurCell
					END IF

		CASE CHR$(0) + "K"                  'LEFT
			currC = (ASC(CurrCell)) - 64
				EraseCursor
				currC = currC - 1
					IF currC < 1 THEN
						currC = 1
						SOUND 3800, 3
					END IF
					IF currC < (FirstX - 64) THEN
						FirstX = FirstX - 1
							CurrCell = CHR$(currC + 64)
							ShowData FirstY, CHR$(FirstX)
							ShowCurCell
					ELSE
						CurrCell = (CHR$(currC + 64))
						ShowLetters CHR$(FirstX)
						ShowNumbers FirstY
						ShowCurCell
					END IF

		CASE CHR$(0) + "G"                  'HOME
			COLOR InfoColor, 0
			LOCATE 25, 40: PRINT "HOME"; : FOR t = 1 TO 15000: NEXT
			FirstY = 1
			FirstX = 65
			CurrCell = "A"
			currRow = 1
			ShowData 1, "A"
			ShowCurCell
			COLOR InfoColor, 0
			LOCATE 25, 40: PRINT "    ";

		CASE CHR$(0) + "O"                  'END
			COLOR InfoColor, 0
			LOCATE 25, 45: PRINT "END"; : FOR t = 1 TO 15000: NEXT
			GetEnd
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell
			COLOR InfoColor, 0
			LOCATE 25, 45: PRINT "   ";

		CASE "/": GOSUB MenuSystem

		CASE "0" TO "9"
			GetNumber
			Changed = 1

		CASE "-"
			GetNumber

		CASE "."
			GetNumber

		CASE "="
			GetFormula

		CASE " " TO "~"
			GetText

	END SELECT
LOOP

END SUB

FUNCTION MMenu (CurrChoiceX, MaxChoice, choice$(), ItemRow(), ItemCol(), Hot$, BarMode)
	COLOR 15, 0
	IF SubMenu = FALSE THEN LOCATE 2, 1: PRINT SPACE$(80);
	ShowStat "MENU"
	currchoice = CurrChoiceX

Start:
	IF BarMode THEN
		COLOR 15, 0
		LOCATE 1, 1
		PRINT SPACE$(76);
	ELSE
		COLOR MenuRegular, 0
		Box ItemRow(1) - 1, ItemCol(1) - 1, ItemRow(MaxChoice) + 1, ItemCol(1) + LEN(choice$(1)), DblBox
			IF SubMenu = FALSE THEN
				IF MenuLeft <> ItemCol(1) - 1 THEN
					LOCATE MenuRow + 1, MenuLeft
						PRINT "*";
				ELSE
					LOCATE MenuRow + 1, MenuLeft
						PRINT "*";
				END IF
				LOCATE MenuRow + 1, MenuRight
					PRINT "*";
			END IF
	END IF

	Howdy = 0
	COLOR MenuRegular, 0
	FOR a = 1 TO MaxChoice
		LOCATE ItemRow(a), ItemCol(a)
		PRINT choice$(a);
			IF MID$(Hot$, a, 1) <> " " THEN
				Letter$ = MID$(Hot$, a, 1)
					FOR t = 1 TO LEN(choice$(a))
						IF MID$(choice$(a), t, 1) = Letter$ THEN
							LOCATE ItemRow(a), ItemCol(a) + t - 1: COLOR MenuHotKey, 0: PRINT Letter$; : COLOR MenuRegular, 0: EXIT FOR
						END IF
					NEXT
			END IF

			IF LEFT$(choice$(a), 1) = "*" THEN
				LOCATE ItemRow(a), ItemCol(a) - 1
					PRINT "*";
				LOCATE ItemRow(a), ItemCol(a) + LEN(choice$(a))
					PRINT "*";
			END IF
	NEXT a

	finished = FALSE

	WHILE NOT finished

		GOSUB MenuShowCursor
		GOSUB MenuGetKey
		GOSUB MenuHideCursor

		SELECT CASE Kbd$
			CASE CHR$(0) + "H": GOSUB MenuUp
			CASE CHR$(0) + "P": GOSUB MenuDown
			CASE CHR$(0) + "K": GOSUB MenuLeft
			CASE CHR$(0) + "M": GOSUB MenuRight
			CASE CHR$(13): GOSUB MenuEnter
			CASE CHR$(27): GOSUB MenuEscape
			CASE " " TO "~": GOSUB GetHotKey
			CASE ELSE: SOUND 3800, 2: SOUND 32767, 1
		END SELECT
	WEND

	IF Howdy = 1 THEN GOTO Start

	MMenu = currchoice
	LOCATE , , 0
	GOSUB EraseHotKeys

	IF ShowOut = 1 THEN
		MenuRow = ItemRow(currchoice)
		LOCATE ItemRow(currchoice), ItemCol(currchoice)
			PRINT "*";
		MenuLeft = ItemCol(currchoice)
		LOCATE ItemRow(currchoice), ItemCol(currchoice) + LEN(choice$(currchoice)) - 1
			PRINT "*";
		MenuRight = ItemCol(currchoice) + LEN(choice$(currchoice)) - 1
	END IF

	EXIT FUNCTION

MenuEnter:
	finished = TRUE
	IF SubMenu = FALSE AND BarMode = TRUE THEN ShowOut = 1
	RETURN

MenuEscape:
	currchoice = 0
	finished = TRUE
	RETURN

MenuUp:
	IF BarMode THEN
		SOUND 3800, 2: SOUND 32767, 1
	ELSE
		currchoice = (currchoice + MaxChoice - 2) MOD MaxChoice + 1
			IF LEFT$(choice$(currchoice), 1) = "*" THEN
				currchoice = (currchoice + MaxChoice - 2) MOD MaxChoice + 1
			END IF
	END IF
	RETURN

MenuLeft:
	IF SubMenu = TRUE THEN GOSUB MenuUp: RETURN

	IF BarMode THEN
		currchoice = (currchoice + MaxChoice - 2) MOD MaxChoice + 1
	ELSE
		currchoice = -2
		finished = TRUE
	END IF
	RETURN

MenuRight:
	IF SubMenu = TRUE THEN GOSUB MenuDown: RETURN

	IF BarMode THEN
		currchoice = (currchoice) MOD MaxChoice + 1
	ELSE
		currchoice = -3
		finished = TRUE
	END IF
	RETURN

MenuDown:
	IF BarMode THEN
		ShowOut = 1
		finished = TRUE
	ELSE
		currchoice = (currchoice) MOD MaxChoice + 1
			IF LEFT$(choice$(currchoice), 1) = "*" THEN
				currchoice = (currchoice) MOD MaxChoice + 1
			END IF
	END IF
	RETURN

MenuShowCursor:
	COLOR 0, MenuHighlight
	LOCATE ItemRow(currchoice), ItemCol(currchoice)
	PRINT choice$(currchoice);
		IF BarMode = FALSE THEN
			IF MID$(Hot$, currchoice, 1) <> " " THEN
				Letter$ = MID$(Hot$, currchoice, 1)
					FOR t = 1 TO LEN(choice$(currchoice))
						IF MID$(choice$(currchoice), t, 1) = Letter$ THEN
							LOCATE ItemRow(currchoice), ItemCol(currchoice) + t - 1: COLOR MenuHotKey, 0: PRINT Letter$; : COLOR 0, MenuHighlight: EXIT FOR
						END IF
					NEXT
			END IF
		END IF
	RETURN

MenuGetKey:
	Kbd$ = ""
	WHILE Kbd$ = ""
		ShowExtra
		Kbd$ = INKEY$
	WEND
	RETURN

MenuHideCursor:
	COLOR MenuRegular, 0
	LOCATE ItemRow(currchoice), ItemCol(currchoice)
	PRINT choice$(currchoice);
			IF MID$(Hot$, currchoice, 1) <> " " THEN
				COLOR 15, 0
				Letter$ = MID$(Hot$, currchoice, 1)
					FOR t = 1 TO LEN(choice$(currchoice))
						IF MID$(choice$(currchoice), t, 1) = Letter$ THEN
							LOCATE ItemRow(currchoice), ItemCol(currchoice) + t - 1: COLOR MenuHotKey, 0: PRINT Letter$; : COLOR MenuRegular, 0: EXIT FOR
						END IF
					NEXT
			END IF
	RETURN

GetHotKey:
	FOR MB = 1 TO MaxChoice
		IF UCASE$(MID$(Hot$, MB, 1)) = UCASE$(Kbd$) THEN
			GOSUB MenuHideCursor
			currchoice = MB
			IF BarMode = FALSE THEN GOSUB MenuShowCursor: FOR t = 1 TO 30000: NEXT: GOSUB MenuHideCursor
			IF SubMenu = FALSE AND BarMode = TRUE THEN ShowOut = 1
			finished = TRUE
		END IF
	NEXT
		IF finished = FALSE THEN SOUND 3800, 2: SOUND 32767, 1
RETURN

EraseHotKeys:
	COLOR MenuRegular, 0
	FOR a = 1 TO MaxChoice
		LOCATE ItemRow(a), ItemCol(a)
		PRINT choice$(a);
			IF LEFT$(choice$(a), 1) = "*" THEN
				LOCATE ItemRow(a), ItemCol(a) - 1
					PRINT "*";
				LOCATE ItemRow(a), ItemCol(a) + LEN(choice$(a))
					PRINT "*";
			END IF
	NEXT a
RETURN
END FUNCTION

SUB N2S (Num$, st$)
	Num$ = LTRIM$(Num$)
	Num$ = RTRIM$(Num$)
	st$ = Num$
END SUB

SUB NestedForm (In$, Out$)
	DIM Num$(100)
	DIM Oper(100) AS STRING * 1
	DIM Letter$(100)
	LParen = 0
	RParen = 0
		FOR t = 1 TO LEN(In$)
			IF MID$(In$, t, 1) = "(" THEN
				LParen = t + 1
			END IF
			IF MID$(In$, t, 1) = ")" AND LParen <> 0 THEN
				RParen = t - 1
			END IF
		NEXT
	IF LParen = 0 THEN Out$ = In$: EXIT SUB

	lft$ = LEFT$(In$, LParen - 2)
	rgt$ = MID$(In$, RParen + 2, LEN(In$))

	NewForm$ = MID$(In$, LParen, RParen - LParen + 1)
FOR It = 1 TO VAL(Iter)

		Range = 0
		FOR t = 1 TO LEN(NewForm$)
			IF MID$(NewForm$, t, 1) = "[" THEN
				Range = Range + 1
			END IF
		NEXT
		IF Range > 0 THEN
			FOR t = 1 TO Range
				FindRange NewForm$, NewForm$
			NEXT
		END IF
				SplitForm NewForm$, Num$(), Num

		FOR t = 1 TO Num
				Letter$(t) = LEFT$(Num$(t), 1)
					IF Letter$(t) = "9" OR Letter$(t) = "8" OR Letter$(t) = "7" OR Letter$(t) = "6" OR Letter$(t) = "5" OR Letter$(t) = "4" OR Letter$(t) = "3" OR Letter$(t) = "2" OR Letter$(t) = "1" OR Letter$(t) = "0" THEN Letter$(t) = " "
			IF RIGHT$(Num$(t), 1) = "+" OR RIGHT$(Num$(t), 1) = "-" OR RIGHT$(Num$(t), 1) = "/" OR RIGHT$(Num$(t), 1) = "*" OR RIGHT$(Num$(t), 1) = "S" THEN
				Oper(t) = RIGHT$(Num$(t), 1)
				IF Letter$(t) <> " " THEN
					Num$(t) = LEFT$(Num$(t), LEN(Num$(t)) - 1)
				ELSE
					Num$(t) = LEFT$(Num$(t), LEN(Num$(t)))
				END IF
			ELSE
				Oper(t) = "+"
			END IF
				Num$(t) = RIGHT$(Num$(t), LEN(Num$(t)) - 1)
		NEXT
				lt = (ASC(Letter$(1)) - 64)
				dt = (VAL(Num$(1)))
				IF Letter$(1) <> " " THEN
					Nxt# = VAL(Cell(lt, dt).CellValue)
				ELSE
					Nxt# = VAL(Num$(1))
				END IF
					TotalValue# = Nxt#
		FOR t = 2 TO Num
				lt = (ASC(Letter$(t)) - 64)
				dt = (VAL(Num$(t)))
				IF Letter$(t) <> " " THEN
					Nxt# = VAL(Cell(lt, dt).CellValue)
				ELSE
					Nxt# = VAL(Num$(t))
				END IF

			IF Oper(t - 1) = "^" THEN TotalValue# = TotalValue# ^ Nxt#
			IF Oper(t - 1) = "*" THEN TotalValue# = TotalValue# * Nxt#
			IF Oper(t - 1) = "/" THEN TotalValue# = TotalValue# / Nxt#
			IF Oper(t - 1) = "+" THEN TotalValue# = TotalValue# + Nxt#
			IF Oper(t - 1) = "-" THEN TotalValue# = TotalValue# - Nxt#
		NEXT
NEXT

		TotalValue$ = STR$(TotalValue#)
		TotalValue$ = RTRIM$(TotalValue$)
		TotalValue$ = LTRIM$(TotalValue$)

Out$ = lft$ + TotalValue$ + rgt$
END SUB

SUB NewSpread
	CellRow = 4
	CellCol = 7
	ShowStat "WAIT"
	FOR col = 1 TO MaxX
		FOR row = 1 TO MaxY
				Cell(col, row).CellValue = ""
				Cell(col, row).Formula = ""
			ShowExtra
		NEXT
	NEXT
	CalcHalves
	FirstX = 65
	FirstY = 1
	CurrCell = "A"
	currRow = 1
	ShowData 1, "A"
	ShowCurCell
	CurrSheet = ""
	NumType = "P" + Decimal
	ShowExtra
END SUB

SUB OpenSpread
		file$ = ".SHT"
		ShowStat "LOAD"
	FileGet file$
		IF file$ = "NO FILES FOUND" THEN
			SOUND 3800, 2
			SOUND 32767, 1
			ShowStat "READY"
			EXIT SUB
		END IF
		IF file$ = "ESC" THEN
			ShowStat "READY"
			EXIT SUB
		END IF
	file$ = UCASE$(file$)
	CALL exist(CurrDrive + ":" + RTRIM$(CurrDir) + "\" + file$ + CHR$(0), exists%)
		IF NOT exists% THEN
			COLOR 15, 0
			Center 1, "Sheet " + CHR$(34) + RTRIM$(file$) + CHR$(34) + " does not exist!"
			SLEEP
			EXIT SUB
		END IF

CurrSheet = UCASE$(file$)
ShowExtra

	OPEN CurrDrive + ":" + RTRIM$(CurrDir) + "\" + CurrSheet FOR INPUT AS #1
		col = 0: row = 0
		FOR col = 1 TO MaxX
			INPUT #1, W(col)
			FOR row = 1 TO MaxY
				LINE INPUT #1, Cell(col, row).Usg
				LINE INPUT #1, Cell(col, row).Formula
				LINE INPUT #1, Cell(col, row).CellValue
					 INPUT #1, Cell(col, row).ID
			NEXT
		NEXT
		INPUT #1, NumType
	CLOSE #1
ShowExtra
CurrCell = "A"
currRow = 1
FirstY = 1
FirstX = 65
ShowData 1, "A"
END SUB

SUB ReCalc
	DIM choice$(3), ItemRow(3), ItemCol(3), scrn(18, 69)  AS INTEGER, colr(18, 69) AS INTEGER

		FOR row = 14 TO 18
			FOR col = 39 TO 69
				scrn(row, col) = SCREEN(row, col)
				colr(row, col) = SCREEN(row, col, 1)
			NEXT
		NEXT

ReCalc:
	SubMenu = TRUE

	choice$(1) = " Mode            " + RIGHT$(SPACE$(10) + RTRIM$(CalcMode), 10) + " " + CHR$(16)
	choice$(2) = " Order          " + RIGHT$(SPACE$(11) + RTRIM$(Order), 11) + " " + CHR$(16)
	choice$(3) = " Iterations             " + RIGHT$(SPACE$(3) + RTRIM$(Iter), 3) + " " + CHR$(16)

	Hot$ = "MOI"

	ItemRow(1) = 15: ItemCol(1) = 40
	ItemRow(2) = 16: ItemCol(2) = 40
	ItemRow(3) = 17: ItemCol(3) = 40

	curr = MMenu(1, 3, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 1: GOSUB NewMode: ClearBuffer: GOTO ReCalc
		CASE 2: GOSUB NewOrder: ClearBuffer: GOTO ReCalc
		CASE 3: GOSUB NewIter: ClearBuffer: GOTO ReCalc
	END SELECT

	SubMenu = FALSE

		FOR row = 14 TO 18
			FOR col = 39 TO 69
				LOCATE row, col
					COLOR colr(row, col), 0
						PRINT CHR$(scrn(row, col));
			NEXT
		NEXT
EXIT SUB

NewMode:
	choice$(1) = " Background  "
	choice$(2) = " Manual      "

	Hot$ = "BM"

	ItemRow(1) = 16: ItemCol(1) = 44
	ItemRow(2) = 17: ItemCol(2) = 44

	curr = MMenu(1, 2, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)
		IF curr <> 0 THEN
			choice$(curr) = LTRIM$(choice$(curr))
			CalcMode = choice$(curr)
		END IF

RETURN
NewOrder:
	choice$(1) = " Natural      "
	choice$(2) = " Column-wise  "
	choice$(3) = " Row-wise     "

	Hot$ = "NCR"

	ItemRow(1) = 15: ItemCol(1) = 47
	ItemRow(2) = 16: ItemCol(2) = 47
	ItemRow(3) = 17: ItemCol(3) = 47

	curr = MMenu(1, 3, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)
		IF curr <> 0 THEN
			choice$(curr) = LTRIM$(choice$(curr))
			Order = choice$(curr)
		END IF

	RETURN
NewIter:
	Box 11, 21, 15, 70, DblBox
		LOCATE 12, 22: PRINT "Number of time to re-evaluate formulas [1..255]:"
			Kbd$ = GetNumeric$(14, 22, RTRIM$(Iter), Iter, 3, 3)
		ShowData FirstY, CHR$(FirstX)
		ShowMenu 4
RETURN
END SUB

SUB SaveINI
	OPEN DefaultDir$ + "\SPREADIT.INI" FOR OUTPUT AS #1
		PRINT #1, MenuHighlight
		PRINT #1, MenuRegular
		PRINT #1, MenuHotKey
		PRINT #1, SheetColor
		PRINT #1, StatColor
		PRINT #1, InfoColor
		PRINT #1, CursorColor
		PRINT #1, DefaultFile
		PRINT #1, CalcMode
		PRINT #1, Order
		PRINT #1, Iter
	CLOSE #1
END SUB

SUB SaveSpread
		DIM scrn(25, 80), colr(25, 80)
	ShowStat "SAVE"
	IF CurrSheet = "" THEN
		file$ = ".SHT"
		FileGet file$
			IF file$ = "ESC" THEN
				file$ = ""
				ShowStat "READY"
				EXIT SUB
			END IF
			IF file$ = "NO FILES FOUND" THEN
				FOR row = 10 TO 14
					FOR col = 15 TO 57
						scrn(row, col) = SCREEN(row, col)
						colr(row, col) = SCREEN(row, col, 1)
					NEXT
				NEXT

				COLOR MenuRegular, 0
					Box 10, 15, 14, 57, DblBox
				LOCATE 11, 16
					PRINT "Enter the name of the file (        .SHT)"
				LOCATE 12, 16
					PRINT "[--------]"
				Kbd$ = GetString$(13, 17, "", fl$, 8, 8)
						IF fl$ <> "" THEN CurrSheet = UCASE$(fl$) + ".SHT"
				FOR row = 10 TO 14
					FOR col = 15 TO 57
						LOCATE row, col
							COLOR colr(row, col)
								PRINT CHR$(scrn(row, col))
					NEXT
				NEXT
					IF fl$ = "" THEN RETURN

			END IF
			IF file$ <> "NO FILES FOUND" THEN CurrSheet = UCASE$(file$)
	END IF
	ShowExtra
	OPEN CurrDrive + ":" + RTRIM$(CurrDir) + "\" + CurrSheet FOR OUTPUT AS #1
		FOR col = 1 TO MaxX
			PRINT #1, W(col)
			FOR row = 1 TO MaxY
				PRINT #1, Cell(col, row).Usg
				PRINT #1, Cell(col, row).Formula
				PRINT #1, Cell(col, row).CellValue
				PRINT #1, Cell(col, row).ID
			NEXT
		NEXT
		PRINT #1, NumType
	CLOSE #1
	GetFreeSpace
	ShowStat "READY"
	Changed = 0
END SUB

SUB SelectFormat
	DIM choice$(6), ItemRow(6), ItemCol(6), scrn(25, 80), colr(25, 80)
SetNumeric:
	SubMenu = TRUE
	choice$(1) = "* Choose Numeric Format *"
	choice$(2) = " Scientific              "
	choice$(3) = " Fixed                   "
	choice$(4) = " Currency                "
	choice$(5) = " Plain                   "
	choice$(6) = " Reset                   "

	Hot$ = " SFCPR"

	ItemRow(1) = 7: ItemCol(1) = 17
	ItemRow(2) = 8: ItemCol(2) = 17
	ItemRow(3) = 9: ItemCol(3) = 17
	ItemRow(4) = 10: ItemCol(4) = 17
	ItemRow(5) = 11: ItemCol(5) = 17
	ItemRow(6) = 12: ItemCol(6) = 17

	curr = MMenu(2, 6, choice$(), ItemRow(), ItemCol(), Hot$, FALSE)

	SELECT CASE curr
		CASE 0
			SubMenu = FALSE
			ShowData FirstY, CHR$(FirstX)
			EXIT SUB
		CASE 2:
			GOSUB GetDecimal
			NumType = "S" + Decimal
			Translate RTRIM$(Cell(ASC(CurrCell) - 64, currRow).CellValue), NewUs$
			Cell(ASC(CurrCell) - 64, currRow).Usg = NewUs$
			Cell(ASC(CurrCell) - 64, currRow).ID = NumType
			ClearBuffer
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell
			RETURN

		CASE 3:
			GOSUB GetDecimal
				NumType = "F" + Decimal
			lt = ASC(CurrCell) - 64
			q = currRow
			Translate RTRIM$(Cell(lt, q).CellValue), NewUs$
			Cell(lt, q).Usg = NewUs$
			Cell(lt, q).ID = NumType
			ClearBuffer
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell
			RETURN
		CASE 4:
			GOSUB Currency

		CASE 5:
			NumType = "P" + Decimal
			lt = ASC(CurrCell) - 64
			q = currRow
			Translate RTRIM$(Cell(lt, q).CellValue), NewUs$
			Cell(lt, q).Usg = NewUs$
			Cell(lt, q).ID = NumType
			ClearBuffer
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell
			RETURN

		CASE 6:
			lt = ASC(CurrCell) - 64
			q = currRow
			Cell(lt, q).ID = "\"
			Cell(lt, q).Usg = ""

	END SELECT

	ClearBuffer
SubMenu = FALSE
ShowData FirstY, CHR$(FirstX)
ShowCurCell
EXIT SUB

Currency:
	LOCATE 1, 1
		GOSUB GetDecimal
					NumType = "C" + Decimal
			IF RTRIM$(Cell(ASC(CurrCell) - 64, currRow).CellValue) <> "" AND LEFT$(Cell(ASC(CurrCell) - 64, currRow).ID, 1) <> "S" THEN
				Translate RTRIM$(Cell(ASC(CurrCell) - 64, currRow).CellValue), NewUs$
				Cell(ASC(CurrCell) - 64, currRow).Usg = NewUs$
				Cell(ASC(CurrCell) - 64, currRow).ID = NumType
			END IF
			ClearBuffer
			ShowData FirstY, CHR$(FirstX)
			ShowCurCell
			ShowStat "READY"
	RETURN

GetDecimal:
		FOR row = 9 TO 13
			FOR col = 23 TO 62
				scrn(row, col) = SCREEN(row, col)
				colr(row, col) = SCREEN(row, col, 1)
			NEXT
		NEXT

			COLOR MenuRegular, 0
			Box 9, 23, 13, 62, DblBox
				LOCATE 10, 24: PRINT "Enter number of decimal places: [0..9]"
					Kbd$ = GetNumeric$(12, 24, Decimal, NewDecimal$, 1, 1)
						IF NewDecimal$ = "" THEN RETURN
							IF VAL(NewDecimal$) > 9 THEN GOTO GetDecimal
						Decimal = NewDecimal$
					ClearBuffer
		FOR row = 9 TO 13
			FOR col = 23 TO 62
				LOCATE row, col
					COLOR colr(row, col), 0
						PRINT CHR$(scrn(row, col))
			NEXT
		NEXT
			RETURN
END SUB

SUB SetFile
	ShowStat "SETFILE"
		COLOR MenuRegular, 0
		Box 8, 15, 12, 43, DblBox
			LOCATE 9, 16: PRINT "Enter the new default file:"
			LOCATE 10, 16: PRINT "[--------]"
					DefaultFile = LEFT$(DefaultFile, LEN(RTRIM$(DefaultFile)) - 4)
				Kbd$ = GetString$(11, 17, RTRIM$(DefaultFile), DF$, 8, 8)
					DF$ = UCASE$(DF$) + ".SHT"
					DefaultFile = DF$
END SUB

SUB SetWidths
	ShowStat "WAIT"
	FOR t = 1 TO MaxX
		W(t) = 9
	NEXT
	FOR lt = 1 TO MaxX
		FOR y = 1 TO MaxY
			IF LEFT$(Cell(lt, y).ID, 1) = "\" THEN
				Cell(lt, y).Usg = ""
			ELSE
				FOR q = 1 TO MaxY
					Cell(lt, q).CellValue = LTRIM$(Cell(lt, q).CellValue)
					IF LEFT$(Cell(lt, q).ID, 1) = "\" THEN
						Cell(lt, q).Usg = ""
					END IF
				NEXT
			END IF
		NEXT
	NEXT
END SUB

SUB ShowCurCell
	CellData$ = ""
	COLOR InfoColor, 0
	LOCATE 25, 8
	crrRow$ = STR$(currRow)
	crrRow$ = RTRIM$(crrRow$)
	crrRow$ = LTRIM$(crrRow$)
	PRINT LEFT$(CurrCell + crrRow$ + STRING$(6, 250), 5);
		lt = ASC(CurrCell) - 64
		dt = currRow
				LOCATE 1, 1
					PRINT SPACE$(75)
				LOCATE 1, 1
					N2S STR$(W(lt)), Wh$
					PRINT "[W"; Wh$; "]"
				LOCATE 2, 1
					PRINT SPACE$(80);
				LOCATE 2, 1
						IF RTRIM$(Cell(lt, dt).Formula) <> "" THEN
							PRINT LEFT$(Cell(lt, dt).Formula, 80)
						ELSEIF RTRIM$(Cell(lt, dt).ID) <> "\" THEN
							PRINT "(" + Cell(lt, dt).ID + ") ";
							PRINT LEFT$(Cell(lt, dt).CellValue, 75)
						ELSEIF RTRIM$(Cell(lt, dt).CellValue) <> "" THEN
							PRINT "'";
							PRINT LEFT$(Cell(lt, dt).CellValue, 79)
						END IF

	LOCATE CellRow, CellCol
	COLOR 0, CursorColor
		FOR col = CellCol TO CellCol + W(lt) - 1
			dta = SCREEN(CellRow, col, 0)
				CellData$ = CellData$ + CHR$(dta)
		NEXT
			PRINT CellData$;
END SUB

SUB ShowData (bgin AS INTEGER, Letter$)
CALL CalcAttr(15, 0, wattr%)

ShowLetters:
IF Clean = 0 THEN ClearSheet
ShowNumbers bgin
ShowLetters Letter$

	COLOR 0, SheetColor
	LOCATE 3, 4

TotalSpaces = 1
lt = (ASC(Letter$)) - 64

PrintData:
	FOR t = 1 TO MaxX
		IF TotalSpaces + W(lt) < 76 THEN
				TotalSpaces = TotalSpaces + W(lt)
				IF lt + 64 > MaxX + 64 THEN EXIT FOR
					cl = POS(0)
					rw = CSRLIN
					IF CHR$(lt + 64) <> CurrCell THEN LOCATE rw, cl + W(lt + 1)
					IF CHR$(lt + 64) = CurrCell THEN LOCATE rw, cl + W(lt + 1)

					pq = 4
					COLOR 15, 0
					FOR dt = bgin TO bgin + 19
						LOCATE pq, cl
							IF dt > MaxY THEN EXIT FOR
							IF RTRIM$(Cell(lt, dt).CellValue) <> "" AND LEFT$(Cell(lt, dt).ID, 1) <> "\" THEN
								IF LEN(RTRIM$(Cell(lt, dt).Usg)) > W(lt) THEN
									PRINT STRING$(W(lt), "*");
										cll = POS(0)
											PRINT LEFT$(SPACE$(80), TT - cll);
								ELSEIF RTRIM$(Cell(lt, dt).CellValue) <> "" THEN
									PRINT USING RIGHT$(SPACE$(W(lt)) + RTRIM$(Cell(lt, dt).Usg), W(lt)); VAL(RTRIM$(Cell(lt, dt).CellValue));
										cll = POS(0)
											PRINT LEFT$(SPACE$(80), TT - cll);
								END IF
							ELSE
								IF RTRIM$(Cell(lt, dt).CellValue) <> "" AND Cell(lt, dt).ID = "\ " THEN
									st$ = LEFT$(Cell(lt, dt).CellValue, TT - cl)
										rt = CSRLIN
										ct = POS(0)
											CALL XQPrintD(st$, rt, ct, wattr%, 0)
								ELSEIF CHR$(FirstX) <> "A" THEN
									FOR col = 1 TO FirstX - 65
										IF LEN(RTRIM$(Cell(col, dt).CellValue)) > W(col) THEN
											GetEndOfLine col, dt, Value$
											LOCATE pq, 4
											IF Value$ <> "" THEN PRINT LEFT$(Value$, TT)
										END IF
									NEXT
								END IF
							END IF
						pq = pq + 1
					NEXT
				COLOR 0, SheetColor
				LOCATE 3, cl + W(lt)
				lt = lt + 1
		END IF
	NEXT
	IF ASC(LastChar) < ASC(CurrCell) THEN
		FirstX = FirstX + (W(lt) / 9)
			lt = FirstX
		Letter$ = CHR$(lt)
		GOTO ShowLetters
	END IF

	COLOR 15, 0
	Clean = 0
END SUB

SUB ShowExtra
	rw = CSRLIN
	cl = POS(0)
		COLOR 0, SheetColor
		LOCATE 24, 4: PRINT "* "; DATE$; " * "; TIME$; " * Sheet:"; LEFT$(CurrSheet + "            ", 12); " * "; CurrDrive; ": "; FreeSpace; " BYTES FREE *";

	IF Check = 1 THEN LOCATE , , 0
		CheckKeys
	IF Check = 1 THEN LOCATE , , 1
	LOCATE rw, cl
END SUB

SUB ShowLetters (Letter$)

CALL CalcAttr(0, SheetColor, attr%)
CALL XQPrintD(SPACE$(76), 3, 4, attr%, 0)

lt = (ASC(Letter$)) - 64

COLOR 0, SheetColor
LOCATE 3, 4

cl = POS(0)
TotalSpaces = 1
FOR t = 1 TO MaxX
	IF TotalSpaces + W(lt) < 76 THEN
		TotalSpaces = TotalSpaces + W(lt)
		IF lt + 64 > MaxX + 64 THEN EXIT FOR
			cl = POS(0)
			IF CHR$(lt + 64) = CurrCell THEN CellCol = POS(0)
			IF CHR$(lt + 64) <> CurrCell THEN PRINT (SPACE$(W(lt) / 2)); CHR$(lt + 64); SPACE$(W(lt) / 2);
			IF CHR$(lt + 64) = CurrCell THEN COLOR SheetColor, 0: PRINT SPACE$(W(lt) / 2); CHR$(lt + 64); : LetterMid = POS(0): PRINT SPACE$(W(lt) / 2); : COLOR 0, SheetColor
			LOCATE 3, cl + W(lt)
			LastChar = CHR$(lt + 64)
			lt = lt + 1
		END IF
NEXT
LOCATE 3, cl + W(lt - 1)
	PRINT SPACE$(77 - TotalSpaces);
LOCATE 3, 4
TT = 80
END SUB

SUB ShowMenu (MenuNumber)
	ClearTop
	COLOR MenuRegular, 0
			LOCATE 1, 1: PRINT " File  Edit  Style  Options  Tools "
	SELECT CASE MenuNumber
	CASE 1
		LOCATE 1, 1: PRINT "*File*"
			Box 2, 1, 14, 31, DblBox
				LOCATE 2, 1: PRINT "******"
				LOCATE 3, 2: PRINT " New Spreadsheet          F7 "
				LOCATE 4, 2: PRINT " Open Spreadsheet         F6" + CHR$(16)
				LOCATE 5, 2: PRINT " Save                 ALT-F6";
					IF CurrSheet = "" THEN
						PRINT CHR$(16)
					ELSE
						PRINT " "
					END IF
				LOCATE 6, 2: PRINT " Save As             CTRL-F6" + CHR$(16)
				LOCATE 7, 2: PRINT " Close              SHIFT-F6 "
				LOCATE 8, 2: PRINT " Erase                       "
				LOCATE 9, 1: PRINT "*****************************͹"
				LOCATE 10, 2: PRINT " DOS Shell            ALT-F7 "
				LOCATE 11, 1: PRINT "*****************************͹"
				LOCATE 12, 2: PRINT " Directory            ALT-F5" + CHR$(16)
				LOCATE 13, 2: PRINT " Exit                     F5" + CHR$(16)
	CASE 2
		LOCATE 1, 7: PRINT "*Edit*"
		Box 2, 7, 5, 21, DblBox
			LOCATE 2, 7: PRINT "******"
				LOCATE 3, 8: PRINT " Insert       "
				LOCATE 4, 8: PRINT " Delete       "
	CASE 3
		LOCATE 1, 13: PRINT "*Style*"
			Box 2, 13, 7, 37, DblBox
				LOCATE 2, 14: PRINT "*******"
				LOCATE 3, 14: PRINT " Numeric Format     F2" + CHR$(16)
				LOCATE 4, 13: PRINT "***********************͹"
				LOCATE 5, 14: PRINT " Column Width       F3" + CHR$(16)
				LOCATE 6, 14: PRINT " Reset Widths   ALT-F3 "
	CASE 4
		LOCATE 1, 20: PRINT "*Options*"
			Box 2, 20, 5, 45, DblBox
				LOCATE 2, 20: PRINT "*********"
				LOCATE 3, 21: PRINT " Colors          ALT-F4" + CHR$(16)
				LOCATE 4, 21: PRINT " Recalculation         " + CHR$(16)
	CASE 5
		LOCATE 1, 29: PRINT "*Tools*"
			Box 2, 29, 4, 47, DblBox
				LOCATE 2, 29: PRINT "*******"
				LOCATE 3, 30: PRINT " Set Defaults   " + CHR$(16)
	CASE 6
		Box 9, 49, 15, 70, DblBox
		LOCATE 10, 50: PRINT " Menu Colors       " + CHR$(16)
		LOCATE 11, 50: PRINT " Sheet Color        "
		LOCATE 12, 50: PRINT " Status Color       "
		LOCATE 13, 50: PRINT " Information Color  "
		LOCATE 14, 50: PRINT " Cursor Color       "
	END SELECT
END SUB

SUB ShowNumbers (bgin AS INTEGER)

CALL CalcAttr(15, 0, wattr%)
	p = 4
		COLOR 0, SheetColor
			FOR t = bgin TO bgin + 19
				LOCATE p, 1
					Num$ = STR$(t)
					N2S Num$, st$
					st$ = LEFT$(st$ + "   ", 3)
						IF t > MaxY THEN PRINT "   ";
						IF t <> currRow AND t <= MaxY THEN
							PRINT st$;
						END IF
						IF t = currRow AND t <= MaxY THEN
							COLOR SheetColor, 0
								PRINT st$;
								CellRow = CSRLIN
							COLOR 0, SheetColor
						END IF
				p = p + 1
			NEXT
END SUB

SUB ShowStat (stat$)
	column = 80 - LEN(stat$) + 1
	COLOR 15, 0
	LOCATE 1, 73: PRINT "       ";
	LOCATE 1, column: COLOR 0, StatColor: PRINT stat$;
	COLOR 15, 0
END SUB

SUB SpinThing (row, col, spins)
spin:
	LOCATE row, col
	PRINT "*"
	FOR t = 1 TO 20000: NEXT
	LOCATE row, col
	PRINT "/"
	FOR t = 1 TO 20000: NEXT
	LOCATE row, col
	PRINT "-"
	FOR t = 1 TO 20000: NEXT
	LOCATE row, col
	PRINT "\"
	FOR t = 1 TO 20000: NEXT
	ShowExtra
spin = spin + 1
	IF spin = spins THEN EXIT SUB
GOTO spin
END SUB

SUB SplitForm (st$, newst$(), s)
	place = 1
	s = 1
	FOR count = 1 TO LEN(st$)
			IF MID$(st$, count, 1) = "+" OR MID$(st$, count, 1) = "-" OR MID$(st$, count, 1) = "*" OR MID$(st$, count, 1) = "/" OR MID$(st$, count, 1) = "S" THEN
				newst$(s) = MID$(st$, place, count - place + 1)
					IF VAL(newst$(s)) <> 0 THEN
						newst$(s) = "0" + newst$(s)
					END IF
				place = count + 1
				s = s + 1
			END IF
	NEXT
	newst$(s) = MID$(st$, place, count - place)
		IF VAL(newst$(s)) <> 0 THEN
			newst$(s) = "0" + newst$(s)
		END IF
END SUB

SUB Translate (Start$, end$)
	Start$ = RTRIM$(Start$)
	nd$ = Start$
	IF LEFT$(Start$, 1) = "." THEN
		Start$ = "0" + Start$
	END IF
	length = VAL(RIGHT$(NumType, 1))
	FOR t = 1 TO LEN(Start$)
		IF MID$(Start$, t, 1) = "." THEN period = 1
	NEXT
	IF period = 0 AND LEFT$(NumType, 1) = "C" OR LEFT$(NumType, 1) = "S" OR LEFT$(NumType, 1) = "F" THEN
		Start$ = Start$ + "." + STRING$(length, "0")
	END IF
	IF LEFT$(NumType, 1) = "C" THEN
		end$ = "$" + STRING$(40, 35)
			IF length <> 0 THEN
				FOR t = 1 TO LEN(Start$)
					IF MID$(Start$, t, 1) = "." THEN
						end$ = LEFT$(end$, t) + "." + STRING$(length, 35)
					END IF
				NEXT
			ELSE
				end$ = "$#"
			END IF
	END IF
	IF LEFT$(NumType, 1) = "S" THEN
		end$ = "##." + STRING$(length, 35) + "^^^^"
	END IF
	IF LEFT$(NumType, 1) = "F" THEN
		end$ = STRING$(LEN(Start$), 35)
			FOR t = 1 TO LEN(Start$)
				IF MID$(Start$, t, 1) = "." THEN
					end$ = LEFT$(end$, t - 1) + "." + STRING$(length, 35)
					period = 1
				END IF
			NEXT
				IF period = 0 THEN end$ = end$ + "." + STRING$(length, 35)
	END IF
	IF LEFT$(NumType, 1) = "P" THEN
		end$ = STRING$(LEN(Start$), 35)
			FOR t = 1 TO LEN(Start$)
				IF MID$(Start$, t, 1) = "." THEN
					end$ = LEFT$(end$, t - 1) + "." + STRING$(LEN(MID$(end$, t + 1)), 35)
				END IF
			NEXT
	END IF
Start$ = nd$
Start$ = LTRIM$(Start$)
IF rt$ <> "" THEN NumType = rt$
END SUB
