# VBA-Projects
Various projects written in Visual Basic

How to use:

IMPORTANT!
Be sure that Excel can run macros on your machine "File->Options->Trust center->Trust center settings->Macro settings->Enable VBA Macros"

/\ "PRETURI MATERIALE" :

  -From the main summary page,go to the desired material page (Corners,Tensioner,Grate,Feronart,Supplies and "Pretui lemn/metal/zincare" are not functional yet)
  -Click on a cell of a valid material name under the "Nume" column (if not nothing will happen) (Example: click on the cell with the text "Ø42.4x2" in the round pipe sheet)

  -To add a new price:
    #  Press Ctrl+P to bring up the Price menu
    #  Type in the new price,change current date (if desired) and press "Adauga pret"
        
  -To calculate a new material:
    #  Press Ctrl+A to bring up the material calculation menu
    #  Depending on the material type the menu elements will change,add the measurements (at least one) with the "Adauga" button
    #  When all the measurements have been added press the "Genereaza formula" button to calculate the Excel formula

  -IMPORTANT! The "Tabla neteda" worksheet:
    # The "Tabla neteda" worksheet creates (if possible) a graphic representation of the way rectangles with "Lungime" Length and "Latime" Width can be arranged on a sheet with "Lungime coala" Length and "Latime tabla" Width
    # Quantities written in the field "Arie" will be ignored (because the shape can be anything) and an appropiate message is displayed to the user
    # In realily metal rectangles can be rotated at various degreees to fit on a metal sheet,but this macro can only rotate them at 90
    # It works for small quantities but,unfortunately,will freeze for larger quantities because the code needs to be optimized/re-written

/\ "Calculator_fire" :

  -Calculates the number of pipe joints,UNP steel beams of "Lungime_fir" length necessary to obtain all the lengths written under the "Lungimi_de_taiat" column
  -For each length,and additional "Grosime panza" (in millimeters) is added,this is the girth of material lost in the cutting process (can be ignored by given the value of 0)
  -Each joint,beam,etc will have a length of "Pierdere capat" (in millimeters) subtracted,this is to compensate for material length imperfections (can be ignored by given the value of 0)
  -If there is more than 1 type of a certain length,the number of lengths can be written under the  "Nr bucati" column
  -If the corresponding row next to a length under the column "Nr bucati" is left empty,that length is considered as having a single unit
  -After all lengths have been introduced press "Calculeaza necesar fire"
  -The result appears at the very bottom but also in the form of a list to help aid the one doing the cutting
  -To erase the lengths press Ctrl+A
  -Results will not be erased but will be overwritten next time the "Calculeaza necesar fire" is pressed

/\ "Schimbare_valori_workbook" :

  IMPORTANT!
  -This file was written specifically to modify the collection of excel price files from the company catalogue,so,for example,it will only modify numeric values 5 columns to the right of a cell containing the text "Indirect expenses",two rows down from a cell containing the text "BNR",etc because all the Excel files in the product catalogue are formatted that way

  -Write or paste a valid filepath next to the "Adresa  fisier workbook de schimbat" cell
  -Write or paste the name/s of the Excel file/s (without the .xls or .xlsx extension) that you wish to modify in the black border below
  -Change the exchange rates,Indirect costs,etc to the desired numeric values
  -Material prices can be changed but an exact name has to be given,if only a fragment is given,then prices of all materials containing said fragment will be modified (ex: writing "42" will changeprices for both "Pipe Ф 42x2 " and "Round Ф 42")
  -When all the desired changes have been written down press the "Schimbare workbook-uri" button
