#pragma once
#include <iostream>
#include <vector>
#include <string>
#include<Windows.h>
#include "node.h"
#include "iterator.h"

using namespace std;

HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
void gotoxy(int x, int y)
{
    COORD coordinates;
    coordinates.X = x;
    coordinates.Y = y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
}

class MiniExcel {

private:
    node* selectedNode;

    void updatePrevCell(node* prev)
    {
        SetConsoleTextAttribute(hConsole, prev->data->getCode());
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4));
        horiBlockLine();
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4) + 1);
        cout << '|' << prev->data->getData() << '|';
        gotoxy((prev->data->getX() * 6), (prev->data->getY() * 4) + 2);
        horiBlockLine();
    }
    void updateSelectedCell()
    {
        SetConsoleTextAttribute(hConsole, selectedNode->data->getCode());
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4));
        horiBlockLine();
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4) + 1);
        cout << '|' << selectedNode->data->getData() << '|';
        gotoxy((selectedNode->data->getX() * 6), (selectedNode->data->getY() * 4) + 2);
        horiBlockLine();
    }

    bool isColumnSame(itrtor start, itrtor end)
    {
        itrtor temp = start;
        while (temp.curr->top != nullptr)
        {
            if (temp.curr == end.curr)
                return true;
            ++temp;
        }
        if (temp.curr == end.curr)
            return true;

        temp = start;
        while (temp.curr->bottom != nullptr)
        {
            if (temp.curr == end.curr)
                return true;
            --temp;
        }
        if (temp.curr == end.curr)
            return true;
        return false;
    }
    bool isRowSame(itrtor start, itrtor end)
    {
        itrtor temp = start;
        bool flag = true;
        while (temp.curr->right != nullptr)
        {
            if (temp.curr == end.curr)
                return flag;
            temp++;
        }
        if (temp.curr == end.curr)
            return flag;

        temp = start;
        while (temp.curr->left != nullptr)
        {
            if (temp.curr == end.curr)
                return flag;
            temp--;
        }
        if (temp.curr == end.curr)
            return flag;
        flag = false;
        return flag;
    }

    void horiBlockLine()
    {
        char horiBlock = 254;
        for (int i = 0; i < 6; i++)
            cout << horiBlock;
    }

    node* getTopLeft()
    {
        node* temp = selectedNode;
        while (temp->top)
        {
            temp = temp->top;
        }
        while (temp->left)
        {
            temp = temp->left;
        }
        return temp;
    }
    node* getTopRight()
    {
        node* temp = selectedNode;
        while (temp->top)
        {
            temp = temp->top;
        }
        while (temp->right)
        {
            temp = temp->right;
        }
        return temp;
    }
    node* getBottomLeft()
    {
        node* temp = selectedNode;
        while (temp->left)
        {
            temp = temp->left;
        }
        while (temp->bottom)
        {
            temp = temp->bottom;
        }
        return temp;
    }
    node* getBottomRight()
    {
        node* temp = selectedNode;
        while (temp->right)
        {
            temp = temp->right;
        }
        while (temp->bottom)
        {
            temp = temp->bottom;
        }
        return temp;
    }
    node* getLeft()
    {
        node* temp = selectedNode;
        while (temp->left)
        {
            temp = temp->left;
        }
        return temp;
    }
    node* getTop()
    {
        node* temp = selectedNode;
        while (temp->top)
        {
            temp = temp->top;
        }
        return temp;
    }

public:
        MiniExcel()
        {
            selectedNode = new node();
            for (int i = 0; i < 4; i++)
            {
                extendColumn();
                extendRow();
            }
            selectedNode->data->select();
        }
        
        void printSheet()
        {
            system("cls");
            node* temp = getTopLeft();
            while (temp)
            {
                node* head = temp;
                while (head)
                {
                    head->location();
                    SetConsoleTextAttribute(hConsole, head->data->getCode());
                    gotoxy((head->data->getX() * 6), (head->data->getY() * 4));
                    horiBlockLine();
                    gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 1);
                    cout << '|' << head->data->getData() << '|';
                    gotoxy((head->data->getX() * 6), (head->data->getY() * 4) + 2);
                    horiBlockLine();
                    
                    head = head->right;
                }
                temp = temp->bottom;
            }
        }
        void updateSheet(node* prev)
        {
            updatePrevCell(prev);
            updateSelectedCell();
        }
        void enterData()
        {
            node* temp = getBottomRight();
            gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
            string input = "";
            cout << "Enter data";
            cin >> input;
            selectedNode->data->setData(input);
            gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
            cout << "                                                                                     ";
            updateSelectedCell();
        }

        void moveUp()
        {   
            node* temp = selectedNode;
            if (selectedNode->top != nullptr)
            {
                selectedNode->data->deselect();
                selectedNode = selectedNode->top;
                selectedNode->data->select();
                updateSheet(temp);
            }
        }
        void moveDown()
        {
            node* temp = selectedNode;
            if (selectedNode->bottom != nullptr)
            {
                selectedNode->data->deselect();
                selectedNode = selectedNode->bottom;
                selectedNode->data->select();
                updateSheet(temp);
            }
            else
            {
                extendRow();
            }
        }
        void moveLeft()
        {
            node* temp = selectedNode;
            if (selectedNode->left != nullptr)
            {
                selectedNode->data->deselect();
                selectedNode = selectedNode->left;
                selectedNode->data->select();
                updateSheet(temp);
            }
        }
        void moveRight()
        {
            node* temp = selectedNode;
            if (selectedNode->right != nullptr)
            {
                selectedNode->data->deselect();
                selectedNode = selectedNode->right;
                selectedNode->data->select();
                updateSheet(temp);
            }
            else
            {
                extendColumn();
            }
        }
        
        void insertRowAbove()
        {
                node* curr = getLeft();
                
                if (curr->top == nullptr)
                {
                    while (curr)
                    {
                        node* newNode = new node();

                        newNode->bottom = curr;
                        curr->top = newNode;

                        curr = curr->right;
                    }
                    curr = getLeft();
                    while (curr->right)
                    {
                        curr->top->right = curr->right->top;
                        curr->top->right->left = curr->top;
                        curr = curr->right;
                    }
                    return;
                }
                while (curr)
                {
                    node* newNode = new node();

                    newNode->top = curr->top;
                    newNode->bottom = curr;
                    curr->top->bottom = newNode;
                    curr->top = newNode;

                    curr = curr->right;
                }
                curr = getLeft();
                while (curr->right)
                {
                    curr->top->right = curr->right->top;
                    curr->top->right->left = curr->top;
                    curr = curr->right;
                }
                printSheet();
        }
        void insertRowBelow()
        {
            node* curr = getLeft();
            if (curr->bottom == nullptr)
            {
                extendRow();
                return;
            }
            while (curr)
            {
                node* newNode = new node();

                newNode->top = curr;
                newNode->bottom = curr->bottom;
                curr->bottom->top = newNode;
                curr->bottom = newNode;

                curr = curr->right;
            }
            curr = getLeft();
            while (curr->right)
            {
                curr->bottom->right = curr->right->bottom;
                curr->bottom->right->left = curr->bottom;
                curr = curr->right;
            }
            printSheet();
        }
        void insertColumnToRight()
        {
            node* curr = getTop();
            if (curr->right == nullptr)
            {
                extendColumn();
                return;
            }
            while (curr)
            {
                node* newNode = new node();

                newNode->left = curr;
                newNode->right = curr->right;
                curr->right->left = newNode;
                curr->right = newNode;

                curr = curr->bottom;
            }
            curr = getTop();
            while (curr->bottom)
            {
                curr->right->bottom = curr->bottom->right;
                curr->right->bottom->top = curr->right;
                curr = curr->bottom;
            }

        }
        void insertColumnToLeft()
        {
            node* curr = getTop();
            if (curr->left == nullptr)
            {
                while (curr)
                {
                    node* newNode = new node();

                    newNode->right = curr;
                    curr->left = newNode;

                    curr = curr->bottom;
                }
                curr = getTop();
                while (curr->bottom)
                {
                    curr->left->bottom = curr->bottom->left;
                    curr->bottom->left->top = curr->left;
                    curr = curr->bottom;
                }
                return;
            }
            while (curr)
            {
                node* newNode = new node();

                newNode->right = curr;
                newNode->left = curr->left;
                curr->left->right = newNode;
                curr->left = newNode;

                curr = curr->bottom;
            }
            curr = getTop();
            while (curr->bottom)
            {
                curr->left->bottom = curr->bottom->left;
                curr->bottom->left->top = curr->left;
                curr = curr->bottom;
            }
        }

        void insertCellByRightShift()
        {
            extendColumn();
            node* tempNode = selectedNode;
            string currData;
            string prevData = "    ";

            while (tempNode)
            {
                currData = tempNode->data->getData();
                tempNode->data->setData(prevData);
                prevData = currData;
                tempNode = tempNode->right;
            }
            printSheet();
        }
        void insertCellByDownShift()
        {
            extendRow();
            node* tempNode = selectedNode;
            string currData;
            string prevData = "    ";

            while (tempNode)
            {
                currData = tempNode->data->getData();
                tempNode->data->setData(prevData);
                prevData = currData;
                tempNode = tempNode->bottom;
            }
            printSheet();
        }

        void deleteCellByLeftShift()
        {
            node* temp = selectedNode;
            moveLeft();
            while (temp->right)
            {
                temp->data->setData(temp->right->data->getData());
                temp = temp->right;
            }
            temp->data->setData("    ");
            printSheet();
        }
        void deleteCellByUpShift()
        {
            node* temp = selectedNode;
            moveUp();
            while (temp->bottom)
            {
                temp->data->setData(temp->bottom->data->getData());
                temp = temp->bottom;
            }
            temp->data->setData("    ");
            printSheet();
        }
        void deleteColumn()
        {
            if (!selectedNode->left && !selectedNode->right)
                return; // Don't delete if there is only one column

            node* temp = getTop();
            node* toBeDeleted = nullptr;

            if (selectedNode->left == nullptr)//if the column to be deleted is most left
                selectedNode = selectedNode->right;
            else                              //if the column to be deleted is most right
                selectedNode = selectedNode->left;

            if (!temp->left)//most left case
                while (temp)
                {
                    temp->right->left = nullptr;
                    toBeDeleted = temp;
                    temp = temp->bottom;
                    delete toBeDeleted;
                }
            else if (!temp->right)//most right case
                while (temp)
                {
                    temp->left->right = nullptr;
                    toBeDeleted = temp;
                    temp = temp->bottom;
                    delete toBeDeleted;
                }
            else//normal Case
                while (temp)
                {
                    temp->left->right = temp->right;
                    temp->right->left = temp->left;
                    temp->left = nullptr;
                    temp->right = nullptr;
                    toBeDeleted = temp;
                    temp = temp->bottom;
                    delete toBeDeleted;
                }
            selectedNode->data->select();
            printSheet();//update the console
        }
        void deleteRow()
        {
            if (!selectedNode->top && !selectedNode->bottom)
                return; // Don't delete if there is only one row

            node* temp = getLeft();
            node* toBeDeleted = nullptr;

            if(selectedNode->top == nullptr)//if the row to be deleted is top row
            selectedNode = selectedNode->bottom;
            else //if(selectedNode->bottom == nullptr)//if the row to be deleted is bottom row
            selectedNode = selectedNode->top;

            if (!temp->top)//top row Case
                while (temp)
                {
                    temp->bottom->top = nullptr;
                    toBeDeleted = temp;
                    temp = temp->right;
                    delete toBeDeleted;
                }
            else if(!temp->bottom)//bottom row case
                while (temp)
                {
                    temp->top->bottom = nullptr;
                    toBeDeleted = temp;
                    temp = temp->right;
                    delete toBeDeleted;
                }
            else//normal Case
                while (temp)
                {
                    temp->bottom->top = temp->top;
                    temp->top->bottom = temp->bottom;
                    temp->top = nullptr;
                    temp->bottom = nullptr;
                    toBeDeleted = temp;   
                    temp = temp->right;
                    delete toBeDeleted;
                }
            
            selectedNode->data->select();
            printSheet();//update the console
        }

        void clearColumn()
        {
            node* curr = getTop();
            string val = "    ";
            while (curr)
            {
                curr->data->setData(val);
                curr = curr->bottom;
            }
            printSheet();
        }
        void clearRow()
        {
            node* curr = getLeft();
            string val = "    ";
            while (curr)
            {
                curr->data->setData(val);
                curr = curr->right;
            }
            printSheet();
        }

        double getRangeSum(itrtor start,itrtor end)
        {
            double sum = 0;
            while (start.curr != end.curr)
            {
                sum += stoi(start.curr->data->getData());
                start++;
            }
            sum += stoi(start.curr->data->getData());
            return sum;
        }
        double getRangeAvg(itrtor start, itrtor end)
        {
            double sum = 0;
            double divider = 1;
            while (start.curr != end.curr)
            {
                sum += stoi(start.curr->data->getData());
                start++;
                divider++;
            }
            sum += stoi(start.curr->data->getData());
            double avg = sum / divider;
            return avg;
        }
        int getRangeCount(itrtor start, itrtor end)
        {
            int count = 1;
            while (start.curr != end.curr)
            {
                start++;
                count++;
            }
            return count;
        }
        double getRangeMin(itrtor start, itrtor end)
        {
            double min = INT_MAX;
            double temp;
            while (start.curr != end.curr)
            {
                temp = stoi(start.curr->data->getData());
                if (min > temp)
                    min = temp;
                start++;
            }
            temp = stoi(start.curr->data->getData());
            if (min > temp)
                min = temp;

            return min;
        }
        double getRangeMax(itrtor start, itrtor end)
        {
            double max = INT_MIN;
            double temp;
            while (start.curr != end.curr)
            {
                temp = stoi(start.curr->data->getData());
                if (max < temp)
                    max = temp;
                start++;
            }
            temp = stoi(start.curr->data->getData());
            if (max < temp)
                max = temp;

            return max;
        }
        
        void copy()
        {}
        void cut()
        {}
        void paste()
        {}

        void extendColumn()
        {
            node* temp = getTopRight();
            while (temp)
            {
                node* newNode = new node();
                temp->right = newNode;
                temp->right->left = temp;
                temp = temp->bottom;
            }
            temp = getTopRight();
            while (temp->left->bottom)
            {
                temp->bottom = temp->left->bottom->right;
                temp->bottom->top = temp;
                temp = temp->bottom;
            }
            printSheet();
        }
        void extendRow()
        {
            node* temp = getBottomLeft();
            while (temp)
            {
                node* newNode = new node();
                temp->bottom = newNode;
                temp->bottom->top = temp;
                temp = temp->right;
            }
            temp = getBottomLeft();
            while (temp->top->right)
            {
                temp->right = temp->top->right->bottom;
                temp->right->left = temp;
                temp = temp->right;
            }
            printSheet();
        }
};