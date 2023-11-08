#pragma once
#include <string>
using namespace std;

enum Color
{
    Red,
    Blue,
    White
};
class Cell
{
    friend class node;
private:
    int x;
    int y;
    Color color;
    string val = "    ";
    string dataType;
public:
    Cell()
    {
        x = y = 0;
        color = White;
        val = "    ";
    }
    Cell(int x, int y, string value, string type)
    {
        x = x;
        y = y;
        val = value;
        color = White;
        dataType = type;
    }
    string getData()
    {
        return val;
    }
    void setData(string value)
    {
        string s = "";
        for (int i = 0; i < 4 && i < value.length(); i++)
        {
            s += value[i];
        }
        if (value.length() < 4)
        {
            for (int i = 0; i < 4 - value.length(); i++)
            {
                s += " ";
            }
        }
        this->val = s;
    }
    int getX()
    {
        return x;
    }
    int getY()
    {
        return y;
    }
    void deselect()
    {
        color = White;
    }
    void select()
    {
        color = Blue;
    }
    int getCode()
    {
        if (color == Red)
            return 4;
        else if (color == Blue)
            return 3;
        else
            return 15;
    }
};

