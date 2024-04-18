#ifndef SORTING_H
#define SORTING_H

#include <QSqlTableModel>

class Interactions {
public:

    void InsertionSortByPrice(QSqlTableModel* model) const {
        int rowCount = model->rowCount();

        for (int i = 1; i < rowCount; ++i) {
            QVariantList currentRowValues;
            int currentPrice = model->data(model->index(i, model->fieldIndex("Ціна путівки"))).toInt();
            int j = i - 1;

            while (j >= 0 && model->data(model->index(j, model->fieldIndex("Ціна путівки"))).toInt() > currentPrice) {
                for (int column = 0; column < model->columnCount(); ++column) {
                    currentRowValues.push_back(model->data(model->index(j, column)));
                    model->setData(model->index(j, column), model->data(model->index(j + 1, column)));
                }
                for (int column = 0; column < model->columnCount(); ++column) {
                    model->setData(model->index(j + 1, column), currentRowValues[column]);
                }
                currentRowValues.clear();
                --j;
            }
        }
    }
};

#endif // SORTING_H
