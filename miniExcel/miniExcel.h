#pragma once
#include <iostream>
#include <fstream>
#include <vector>
#include <string>
#include <tuple>
#include<Windows.h>
#include "node.h"
#include "iterator.h"
#include <sstream>
#include <regex>

using namespace std;

HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
void gotoxy(int x, int y)
{
    COORD coordinates;
    coordinates.X = x;
    coordinates.Y = y;
    SetConsoleCursorPosition(GetStdHandle(STD_OUTPUT_HANDLE), coordinates);
}
bool isInteger(string val)
{
    istringstream str(val);
    int num;
    str >> num;
    return str.eof() && !str.fail();
}
bool isFloat(string val)
{
    istringstream str(val);
    float num;
    str >> num;
    return str.eof() && !str.fail();
}
string removeSpaces(string s)
{
    regex pattern("^\\s+|\\s+$");
    return regex_replace(s, pattern, "");
}

class MiniExcel {

private:

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
    void printVector(vector<vector<string>> arr)
    {
        for (int i = 0; i < arr.size(); i++)
        {
            for (int j = 0; j < arr[i].size(); j++)
            {
                cout << arr[i][j] << '\t';
            }
            cout << endl;
        }
    }
    tuple<itrtor, itrtor> itratorAlignment(itrtor start, itrtor end)
    {
        itrtor st = start;

        int startRow = start.curr->data->getY();
        int startCol = start.curr->data->getX();
        int endRow = end.curr->data->getY();
        int endCol = end.curr->data->getX();

        if (startRow > endRow && startCol > endCol)
        {
            st = end;
            end = start;
        }
        else if (startRow > endRow && startCol < endCol)
        {
            while (start.curr->data->getY() != endRow)
                ++start;//move up
            while (end.curr->data->getY() != startRow)
                --end;//move down
            st = start;
        }
        else if (startRow < endRow && startCol > endCol)
        {
            while (start.curr->data->getY() != endRow)
                --start;//move down
            while (end.curr->data->getY() != startRow)
                ++end;//move up
            st = end;
            end = start;
        }
        return make_tuple(st, end);
    }
    tuple<int, int> getRowColCount(string filename)
    {
        ifstream file(filename);
        if (!file.is_open()) {
            cerr << "Unable to open file" << endl;
            return make_tuple(-1, -1);
        }

        char ch = 255;
        int rowCount = 0;
        int colCount = 0;
        string line;
        while (getline(file, line))
        {
            rowCount++;
            if (rowCount == 1)
            {
                istringstream str(line);
                string token;
                while (getline(str, token, ch))
                    colCount++;
            }
        }
        file.close();
        return make_tuple(rowCount, colCount);
    }

    node* getCell(int col, int row)
    {
        node* temp = getTopLeft();
        for (int i = 0; i < col; i++)
            temp = temp->right;
        
        for (int i = 0; i < row; i++)
            temp = temp->bottom;
        
        return temp;
    }

    void horiBlockLine()
    {
        char horiBlock = 254;
        for (int i = 0; i < 6; i++)
            cout << horiBlock;
    }
    void updateSheet(node* prev)
    {
        updatePrevCell(prev);
        updateSelectedCell();
    }

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

public:
        node* selectedNode;
        MiniExcel()
        {
            selectedNode = new node();
            for (int i = 0; i < 4; i++)
                extendColumn();
            for(int i = 0 ;i < 4; i++)
                extendRow();
          
            selectedNode->data->select();
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


        void printSheet()
        {
            system("cls");
            node* temp = getTopLeft();
            while (temp)
            {
                node* head = temp;
                while (head)
                {
                    head->setLocation();
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
        void enterData()
        {
            node* temp = getBottomRight();
            gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
            string input = "";
            cout << "Enter data : ";
            getline(cin,input);
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
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);

            int startRow = st.curr->data->getY();
            int startCol = st.curr->data->getX();
            int endRow = en.curr->data->getY();
            int endCol = en.curr->data->getX();
            const int noOfRows = abs(startRow - endRow);
            const int noOfCols = abs(startCol - endCol);
            double sum = 0;

            for (int i = 0; i <= noOfRows; i++)
            {
                for (int j = 0; j <= noOfCols; j++)
                {
                    if(st.curr->data->getDataType() != String )
                        sum += stof(st.curr->data->getData());
                    else
                    {
                        node* temp = getBottomRight();
                        gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
                        cout << "some of the selected cells does not have numeric data";
                    }
                    st++;//move Right
                }
                while (st.curr->data->getX() != startCol)
                    st--;//move left
                --st;//move down
            }

            return sum;
        }
        double getRangeAvg(itrtor start, itrtor end)
        {
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);

            int startRow = st.curr->data->getY();
            int startCol = st.curr->data->getX();
            int endRow = en.curr->data->getY();
            int endCol = en.curr->data->getX();
            const int noOfRows = abs(startRow - endRow);
            const int noOfCols = abs(startCol - endCol);
            double sum = 0;
            int divider = 0;

            for (int i = 0; i <= noOfRows; i++)
            {
                for (int j = 0; j <= noOfCols; j++)
                {
                    if (st.curr->data->getDataType() != String)
                    {
                        sum += stof(st.curr->data->getData());
                        divider++;
                    }
                    else
                    {
                        node* temp = getBottomRight();
                        gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
                        cout << "some of the selected cells does not have numeric data";
                    }
                    st++;//move Right
                }
                while (st.curr->data->getX() != startCol)
                    st--;//move left
                --st;//move down
            }

            return sum/divider;
        }
        int getRangeCount(itrtor start, itrtor end)
        {
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);

            int startRow = st.curr->data->getY();
            int startCol = st.curr->data->getX();
            int endRow = en.curr->data->getY();
            int endCol = en.curr->data->getX();
            const int noOfRows = abs(startRow - endRow)+1;
            const int noOfCols = abs(startCol - endCol)+1;

            return noOfRows*noOfCols;
        }
        double getRangeMin(itrtor start, itrtor end)
        {
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);

            int startRow = st.curr->data->getY();
            int startCol = st.curr->data->getX();
            int endRow = en.curr->data->getY();
            int endCol = en.curr->data->getX();
            const int noOfRows = abs(startRow - endRow);
            const int noOfCols = abs(startCol - endCol);
            double min = INT_MAX;

            for (int i = 0; i <= noOfRows; i++)
            {
                for (int j = 0; j <= noOfCols; j++)
                {
                    if (st.curr->data->getDataType() != String)
                    {
                        if (min > stof(st.curr->data->getData()))
                            min = stof(st.curr->data->getData());
                    }
                    else
                    {
                        node* temp = getBottomRight();
                        gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
                        cout << "some of the selected cells does not have numeric data";
                    }
                    st++;//move Right
                }
                while (st.curr->data->getX() != startCol)
                    st--;//move left
                --st;//move down
            }

            return min;
        }
        double getRangeMax(itrtor start, itrtor end)
        {
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);

            int startRow = st.curr->data->getY();
            int startCol = st.curr->data->getX();
            int endRow = en.curr->data->getY();
            int endCol = en.curr->data->getX();
            const int noOfRows = abs(startRow - endRow);
            const int noOfCols = abs(startCol - endCol);
            double max = INT_MIN;

            for (int i = 0; i <= noOfRows; i++)
            {
                for (int j = 0; j <= noOfCols; j++)
                {
                    if (st.curr->data->getDataType() != String)
                    {
                        if (max < stof(st.curr->data->getData()))
                            max = stof(st.curr->data->getData());
                    }
                    else
                    {
                        node* temp = getBottomRight();
                        gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
                        cout << "some of the selected cells does not have numeric data";
                    }
                    st++;//move Right
                }
                while (st.curr->data->getX() != startCol)
                    st--;//move left
                --st;//move down
            }

            return max;
        }
        
        vector<vector <string>> copy(itrtor start,itrtor end)
        {
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);
                
                int startRow = st.curr->data->getY();
                int startCol = st.curr->data->getX();
                int endRow = en.curr->data->getY();
                int endCol = en.curr->data->getX();
                const int noOfRows = abs(startRow - endRow);
                const int noOfCols = abs(startCol - endCol);
              
           vector<vector <string>> arr(noOfRows + 1,vector<string>(noOfCols + 1));

           for (int i = 0; i <= noOfRows ; i++)
                {
                    for (int j = 0; j <= noOfCols; j++)
                    {
                        arr[i][j] = st.curr->data->getData();
                            st++;//move Right
                    }
                    while (st.curr->data->getX() != startCol)
                        st--;//move left
                    --st;//move down
                }
           node* temp = getBottomRight();
           gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
           cout << "copied";
                return arr;
        }
        void paste(vector<vector<string>> arr)
        {
            itrtor start = itrtor(selectedNode);
            int Row = start.curr->data->getY();
            int Col = start.curr->data->getX();

            for (int i = 0; i < arr.size(); ++i)
            {
                for (int j = 0; j < arr[i].size(); ++j) 
                {
                    start.curr->data->setData(arr[i][j]);

                    if (start.curr->right == nullptr && j < arr[i].size()-1)
                        extendColumn();

                    start++; // move Right
                }
                while (start.curr->data->getX() != Col)
                    start--;
                if (start.curr->bottom == nullptr && i <arr.size()-1)
                    extendRow();

                --start; // move down
            }
            printSheet();
        }
        vector<vector <string>> cut(itrtor start,itrtor end)
        {
            auto itr = itratorAlignment(start, end);
            itrtor st = get<0>(itr);
            itrtor en = get<1>(itr);

            int startRow = st.curr->data->getY();
            int startCol = st.curr->data->getX();
            int endRow = en.curr->data->getY();
            int endCol = en.curr->data->getX();
            const int noOfRows = abs(startRow - endRow);
            const int noOfCols = abs(startCol - endCol);

            vector<vector <string>> arr(noOfRows + 1, vector<string>(noOfCols + 1));

            for (int i = 0; i <= noOfRows; i++)
            {
                for (int j = 0; j <= noOfCols; j++)
                {
                    arr[i][j] = st.curr->data->getData();
                    st.curr->data->setData("");
                    st++;//move Right
                }
                while (st.curr->data->getX() != startCol)
                    st--;//move left
                --st;//move down
            }
            node* temp = getBottomRight();
            gotoxy(temp->data->getX() * 6, (temp->data->getY() * 4) + 3);
            cout << "copied";
            printSheet();
            return arr;
        }

        void saveDataToFile(string filename)
        {
            char ch = 255;
            ofstream file(filename);
            if (!file.is_open())
                return;
            
            node* temp = getTopLeft();
            while (temp != nullptr)
            {
                node* head = temp;

                while (head != nullptr)
                {
                    file << head->data->getData() << ch;

                    head = head->right;
                }
                file << "\n";
                temp = temp->bottom;
            }

            file.close();
        }
        void loadDataFromFile(string filename)
        {
            char ch = 255;
            int row = 0;
            int col = 0;
            auto count = getRowColCount(filename);
            int rowCount = get<0>(count);
            int colCount = get<1>(count);

            node* temp = getTopLeft();
            while (temp->right)
            {
                col++;
                temp = temp->right;
            }
            while (temp->bottom)
            {
                row++;
                temp = temp->bottom;
            }
            while (row < rowCount-1)
            {
                extendRow();
                row++;
            }
            while (col < colCount-1)
            {
                extendColumn();
                col++;
            }

            ifstream file(filename);
            if (!file.is_open())
                return;
            
            string line;
            row = 0;
            while (getline(file, line))
            {
                istringstream str(line);
                string token;
                col = 0;
                while (getline(str, token, ch))
                {
                    node* currentCell = this->getCell(col, row);

                    if (currentCell != nullptr)
                    {
                        currentCell->data->setData(removeSpaces(token));
                    }
                    col++;
                }
                row++;
            }
            file.close();
            printSheet();
        }
};