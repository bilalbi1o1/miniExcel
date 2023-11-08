#pragma once
#include "Cell.h"

class node
{
    friend class MiniExcel;

public:
    Cell* data;
    node* top;
    node* bottom;
    node* left;
    node* right;
    node()
    {
        data = new Cell();
        top = nullptr;
        bottom = nullptr;
        left = nullptr;
        right = nullptr;
    }
    node(Cell* value)
    {
        data = value;
        top = nullptr;
        bottom = nullptr;
        left = nullptr;
        right = nullptr;
    }
    void location()
    {
        node* newNode = this;
        int counter = 0;
        while (newNode->top != nullptr)
        {
            counter++;
            newNode = newNode->top;
        }
        data->y = counter;
        counter = 0;
        while (newNode->left != nullptr)
        {
            counter++;
            newNode = newNode->left;
        }
        data->x = counter;
    }
};
