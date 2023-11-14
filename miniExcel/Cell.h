#pragma once
#include <string>
//#include <sstream>

using namespace std;

enum Color
{
    Red,
    Blue,
    White
};
enum dataType
{
    String,
    Int,
    Float
};
bool isInteger(string val);
bool isFloat(string val);


class Cell
{
    friend class node;
private:
    int x;//col
    int y;//row
    Color color;
    string val = "    ";

public:
    dataType type;
    Cell()
    {
        x = y = 0;
        color = White;
        val = "    ";
    }
    Cell(int x, int y, string value,dataType type)
    {
        x = x;
        y = y;
        val = value;
        color = White;
        type = String;
    }
    string getData()
    {
        return val;
    }
    void setData(string value)
    {
     
        setDataType(value);

        string str = "";
        for (int i = 0; i < 4 && i < value.length(); i++)
        {
            str += value[i];
        }
        if (value.length() < 4)
        {
            for (int i = 0; i < 4 - value.length(); i++)
            {
                str += " ";
            }
        }
        this->val = str;
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
    dataType getDataType()
    {
        return type;
    }
    void getTypeCode(int code)
    {
        if (code == 1)
            type = Int;
        else if (code == 2)
            type = Float;
        else
            type = String;
    }
    void setDataType(string value)
    {
        if (isInteger(value))
            getTypeCode(1);
        else if (isFloat(value))
            getTypeCode(2);
        else
            getTypeCode(3);
    }

};

