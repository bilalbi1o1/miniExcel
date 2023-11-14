#include <iostream>
#include <conio.h>
#include "miniExcel.h"

using namespace std;

void menu();
int main()
{
    while (true)
    {
        system("cls");
        int choice;
        
        cout << "1 ExcelSheet" << endl;
        cout << "2 View Controls" << endl;
        cout << "3 Exit" << endl;
        cout << "Enter your choice: ";
        cin >> choice;

        if (choice == 1)
        {
            MiniExcel excel;
            itrtor itr1 = excel.getTopLeft();
            itrtor itr2 = excel.getTopLeft();
            vector<vector <string>> arr;
            excel.printSheet();

            while (true)
            {

                if (GetAsyncKeyState(VK_LEFT))
                    excel.moveLeft();
                if (GetAsyncKeyState(VK_DOWN))
                    excel.moveDown();
                if (GetAsyncKeyState(VK_UP))
                    excel.moveUp();
                if (GetAsyncKeyState(VK_RIGHT))
                    excel.moveRight();


                if (GetAsyncKeyState(VK_F1))//for row above
                    excel.insertRowAbove();
                if (GetAsyncKeyState(VK_F2))//for row below
                    excel.insertRowBelow();
                if (GetAsyncKeyState(VK_F3))//for col left
                    excel.insertColumnToLeft();
                if (GetAsyncKeyState(VK_F4))//for col right
                    excel.insertColumnToRight();


                if (GetAsyncKeyState(VK_F5))//to insert data by right shift
                    excel.insertCellByRightShift();
                if (GetAsyncKeyState(VK_F6))//to insert data by down shift
                    excel.insertCellByDownShift();


                if (GetAsyncKeyState(VK_F7))//to delete data by left shift
                    excel.deleteCellByLeftShift();
                if (GetAsyncKeyState(VK_F8))//to delete data by up shift
                    excel.deleteCellByUpShift();
                if (GetAsyncKeyState(VK_F9))//to delete the selected column
                    excel.deleteColumn();
                if (GetAsyncKeyState(VK_F10))//to delete the selected row
                    excel.deleteRow();


                if (GetAsyncKeyState(VK_F11))//to clear the selected column
                    excel.clearColumn();
                if (GetAsyncKeyState(VK_F12))//to clear the selected row
                    excel.clearRow();


                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('9'))//to Save data in the file
                    excel.saveDataToFile("excelData.txt");
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('0'))//to load data from the file
                    excel.loadDataFromFile("excelData.txt");


                if (GetAsyncKeyState(VK_SHIFT))//for Input In Selected Cell
                    excel.enterData();
                if (GetAsyncKeyState(VK_OEM_MINUS))//to select the first iterator
                {
                    itr1 = excel.selectedNode;
                    cout << "selected1";
                } 
                if (GetAsyncKeyState(VK_OEM_PLUS))//to select the second iterator
                {
                    itr2 = excel.selectedNode;
                    cout << "selected2";
                }
                if ((GetAsyncKeyState(VK_CONTROL) & 0x8000) && (GetAsyncKeyState('6') & 0x8000))//Copy
                    arr = excel.copy(itr1,itr2);
                if ((GetAsyncKeyState(VK_CONTROL) & 0x8000) && (GetAsyncKeyState('7') & 0x8000))//Cut
                    arr = excel.cut(itr1, itr2);
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('8'))//Paste
                {
                    if(arr.size() != 0)
                        excel.paste(arr);              
                }      
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('4'))//range Sum
                {
                    double sum = excel.getRangeSum(itr1, itr2);
                    excel.selectedNode->data->setData(to_string(sum));
                    excel.updateSelectedCell();
                }
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('5'))//range Avg
                {
                    double avg = excel.getRangeAvg(itr1, itr2);
                    excel.selectedNode->data->setData(to_string(avg));
                    excel.updateSelectedCell();
                } 
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('1'))//range Count
                {
                    int count = excel.getRangeCount(itr1, itr2);
                    excel.selectedNode->data->setData(to_string(count));
                    excel.updateSelectedCell();
                }   
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('2'))//range Min
                {
                    double min = excel.getRangeMin(itr1, itr2);
                    excel.selectedNode->data->setData(to_string(min));
                    excel.updateSelectedCell();
                }   
                if (GetAsyncKeyState(VK_CONTROL) && GetAsyncKeyState('3'))//range Max
                {
                    double max = excel.getRangeMax(itr1, itr2);
                    excel.selectedNode->data->setData(to_string(max));
                    excel.updateSelectedCell();
                }   
                if (GetAsyncKeyState(VK_ESCAPE))
                    break;

                        Sleep(90);
            }
        }
        else if (choice == 2)
        {
            menu();
        }
        else if (choice == 3)
        {
            break;
        }
        else
        {
            cout << "invalid Choice!!";
            _getch();
        }
    }
}
char getCharAtxy(short int x, short int y)
{
    CHAR_INFO ci;
    COORD xy = { 0, 0 };
    SMALL_RECT rect = { x, y, x, y };
    COORD coordBufSize;
    coordBufSize.X = 1;
    coordBufSize.Y = 1;
    return ReadConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), &ci, coordBufSize, xy, &rect) ? ci.Char.AsciiChar : 'B';
}
void menu()
{
    system("cls");
    SetConsoleTextAttribute(hConsole,4);
    cout << "                        .......................Instructions....................                              " << endl;
    SetConsoleTextAttribute(hConsole,3);
    cout << "1. Up key to move up" << endl;
    cout << "2. Down key to move down" << endl;
    cout << "3. Left key to move left" << endl;
    cout << "4. Right key to move right" << endl;
    cout << "5. F1 to insert row above current cell" << endl;
    cout << "6. F2 to insert row below current cell" << endl;
    cout << "7. F3 to insert column to the left of the current cell" << endl;
    cout << "8. F4 to insert column to the right of the current cell" << endl;
    cout << "9. F5 to insert data by right shift" << endl;
    cout << "10. F6 to insert data by down shift" << endl;
    cout << "11. F7 to delete data by left shift" << endl;
    cout << "12. F8 to delete data by down shift" << endl;
    cout << "13. F9 to delete column of selected cell" << endl;
    cout << "14. F10 to delete row of selected cell" << endl;
    cout << "15. F11 to clear the column of the selected cell" << endl;
    cout << "16. F12 to clear the row of the selected cell" << endl;
    cout << "17. Ctrl + 9 to save data in the file" << endl;
    cout << "18. Ctrl + 0 to load data from the file" << endl;
    cout << "19. Shift to enter data in a cell" << endl;
    cout << "20. Minus Sign to select starting cell of range" << endl;
    cout << "21. Plus Sign to select ending cell of range" << endl;
    cout << "22. Ctrl + 6 for copy operation after selecting its range" << endl;
    cout << "23. Ctrl + 8 for paste operation after selecting its range" << endl;
    cout << "24. Ctrl + 7 for cut operation after selecting its range" << endl;
    cout << "25. Ctrl + 4 for sum operation after selecting its range" << endl;
    cout << "26. Ctrl + 5 for average operation after selecting its range" << endl;
    cout << "27. Ctrl + 1 for count operation after selecting its range" << endl;
    cout << "28. Ctrl + 2 for minimum operation after selecting its range" << endl;
    cout << "29. Ctrl + 3 for maximum operation after selecting its range" << endl;
    cout << "30. Escape to return to main screen" << endl;
    SetConsoleTextAttribute(hConsole, 13);
    cout << "Enter any key to continue...";
    SetConsoleTextAttribute(hConsole, 15);
    _getch();


}