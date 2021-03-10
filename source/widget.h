#pragma once

#include <QWidget>
#include <QDir>

class QLineEdit;
class QTextEdit;

// ======================================================================
class FileFinder : public QWidget {
    Q_OBJECT
private:
    QLineEdit* m_ptxtDir;
    QLineEdit* m_ptxtMask;
    QTextEdit* m_ptxtResult;

public:
    FileFinder(QWidget* pwgt = nullptr);

    void start(const QDir& dir);

public slots:
    void slotFind  ();
    void slotBrowse();
    void slotDocx();

};

