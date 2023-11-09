#pragma once
#include "node.h"

class itrtor
{
    friend class node;
public:
    node* curr;
    itrtor(node* itr)
    {
        curr = itr;
    }
    bool operator!=(itrtor* itr)
    {
        return curr != itr->curr;
    }
    bool operator==(itrtor* itr)
    {
        return curr == itr->curr;
    }
    Cell*& operator*()
    {
        return curr->data;
    }
    itrtor& operator++(int)
    {
        if (curr && curr->right)
            curr = curr->right;

        return *this;
    }
    itrtor& operator--(int)
    {
        if (curr && curr->left)
            curr = curr->left;

        return *this;
    }
    itrtor& operator++()
    {
        if (curr && curr->top)
            curr = curr->top;

        return *this;
    }
    itrtor& operator--()
    {
        if (curr && curr->bottom)
            curr = curr->bottom;

        return *this;
    }
};
