#include <iostream>
#include "miniExcel.h"

using namespace std;

int main()
{
    MiniExcel excel;
    itrtor itr1 = itrtor(excel.getTopLeft());
    itrtor itr2 = itrtor(excel.getTopRight());
    excel.printSheet();
    while (true)
    {

        if (GetAsyncKeyState(VK_SPACE))
        {
            excel.enterData();
        }
        if (GetAsyncKeyState(VK_LEFT))
        {
            excel.moveLeft();
        }
        if (GetAsyncKeyState(VK_DOWN))
        {
            excel.moveDown();
        }
        if (GetAsyncKeyState(VK_UP))
        {
            excel.moveUp();
        }
        if (GetAsyncKeyState(VK_RIGHT))
        {
            excel.moveRight();
        }
        if (GetAsyncKeyState(VK_SHIFT))
        {
            excel.extendRow();
        }
        if (GetAsyncKeyState(VK_TAB))
        {
            excel.extendColumn();
        }
        if (GetAsyncKeyState(VK_F1))
        {
            cout << excel.isRowSame(itr1,itr2);
            cout << excel.isColumnSame(itr1,itr2);
        }
        if (GetAsyncKeyState(VK_F2))
        {
            excel.insertRowAbove();
        }
        if (GetAsyncKeyState(VK_F3))
        {
            excel.insertRowBelow();
        }
        if (GetAsyncKeyState(VK_F4))
        {
            double sum = excel.getRangeMax(itr1,itr2);
            cout << sum;
        }
        if (GetAsyncKeyState(VK_F5))
        {
            excel.clearColumn();
        }
        if (GetAsyncKeyState(VK_F6))
        {
            excel.clearRow();
        }
        Sleep(90);
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