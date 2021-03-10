#include <QtWidgets>
#include "widget.h"
#include "ActiveQt/ActiveQt"
#include <QAxObject>


// ----------------------------------------------------------------------
FileFinder::FileFinder(QWidget* pwgt/*= 0*/) : QWidget(pwgt)
{
    m_ptxtDir    = new QLineEdit(QDir::current().absolutePath());
    m_ptxtMask   = new QLineEdit("*.cpp *.c *.h");
    m_ptxtResult = new QTextEdit;

    QLabel*      plblDir  = new QLabel("&Directory");
    QLabel*      plblMask = new QLabel("&Mask");
    QPushButton* pcmdDir  = new QPushButton("...");
    QPushButton* pcmdFind = new QPushButton("&Find");
    QPushButton* pcmdDocx = new QPushButton("&MakeFile");

    connect(pcmdDir, SIGNAL(clicked()), SLOT(slotBrowse()));
    connect(pcmdFind, SIGNAL(clicked()), SLOT(slotFind()));
    connect(pcmdDocx, SIGNAL(clicked()), SLOT(slotDocx()));

    plblDir->setBuddy(m_ptxtDir);
    plblMask->setBuddy(m_ptxtMask);

    //Layout setup
    QGridLayout* pgrdLayout = new QGridLayout;
    pgrdLayout->setMargin(5);
    pgrdLayout->setSpacing(15);
    pgrdLayout->addWidget(plblDir, 0, 0);
    pgrdLayout->addWidget(plblMask, 1, 0);
    pgrdLayout->addWidget(m_ptxtDir, 0, 1);
    pgrdLayout->addWidget(m_ptxtMask, 1, 1);
    pgrdLayout->addWidget(pcmdDir, 0, 2);
    pgrdLayout->addWidget(pcmdFind, 1, 2);
    pgrdLayout->addWidget(pcmdDocx, 2, 2);
    pgrdLayout->addWidget(m_ptxtResult, 3, 0, 1, 3);
    setLayout(pgrdLayout);
}

// ----------------------------------------------------------------------
void FileFinder::slotBrowse()
{
    QString str = QFileDialog::getExistingDirectory(nullptr,
                                                    "Select a Directory",
                                                    m_ptxtDir->text()
                                                   );

    if (!str.isEmpty()) {
        m_ptxtDir->setText(str);
    }
}

// ----------------------------------------------------------------------
void FileFinder::slotFind()
{
    m_ptxtResult->clear();
    start(QDir(m_ptxtDir->text()));
}
// ----------------------------------------------------------------------
void FileFinder::start(const QDir& dir)
{
    QApplication::processEvents();

    QStringList listFiles =
        dir.entryList(m_ptxtMask->text().split(" "), QDir::Files);

    foreach (QString file, listFiles) {
        m_ptxtResult->append(dir.relativeFilePath(file));
    }

    QStringList listDir = dir.entryList(QDir::Dirs);
    foreach (QString subdir, listDir) {
        if (subdir == "." || subdir == "..") {
            continue;
        }
        start(QDir(dir.absoluteFilePath(subdir)));
    }
}
// ----------------------------------------------------------------------
void FileFinder::slotDocx()
{
    QDir dir(m_ptxtDir->text());

    //открываем файл .docx
    QAxObject* WordApplication = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord
    QAxObject* WordDocuments = WordApplication->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов"
    QAxObject* titleDoc = WordDocuments->querySubObject( "Open(const QString&)", QDir::currentPath() + "/Title1.docx" ); // Открываю документ с титульником

    // отключение грамматики
    QAxObject* Grammatic = WordApplication->querySubObject("Options()");
    Grammatic->setProperty("CheckSpellingAsYouType(bool)", false); // отключение грамматики
    QAxObject* ActiveDocument = WordApplication->querySubObject("ActiveDocument()");

    // нумерация страниц
    ActiveDocument->querySubObject("Sections(Index)", 1)
            ->querySubObject("Footers(WdHeaderFooterIndex)", "wdHeaderFooterPrimary")
            ->querySubObject("PageNumbers")
            ->dynamicCall("Add(PageNumberAlignment, FirstPage)", TRUE, FALSE); // расположение по середине, начало со 2 страницы

    //оформление
    QAxObject* Range = ActiveDocument->querySubObject("Range()");
    Range->setProperty("Alignment", 1); //выравнивание

    QAxObject *font = Range->querySubObject("Font()");
    font->setProperty("Size", 14); //размер заголовка
    font->setProperty("Bold", 0); //не жирный
    font->setProperty("Name", "Times New Roman"); // шрифт
    font->setProperty("Spacing", 0); //интервал между символами

    QAxObject *paragraph = Range->querySubObject("ParagraphFormat()");
    paragraph->setProperty("LineSpacing", 15); //межстрочный интервал

        QDirIterator it(
                    m_ptxtDir->text(),
                    QStringList() << m_ptxtMask->text().split(" "),
                    QDir::Files,
                    QDirIterator::Subdirectories
                    );

        QString previousDirname;
        int i = 1;
        int j = 0;

        while (it.hasNext())
        {
            QFile file(it.next());
            QFileInfo fi(file);

            QString dirname = fi.absolutePath().right(fi.absolutePath().size() - fi.absolutePath().lastIndexOf(QChar('/')) - 1);
            QString filename = fi.fileName();
            if ( file.open( QIODevice::ReadOnly | QIODevice::Text) )
            {
                QString text = file.readAll();

                if (dirname == previousDirname)
                {
                    j++;                    

                    Range->dynamicCall("InsertAfter(QString)", i - 1);
                    Range->dynamicCall("InsertAfter(QString)", ".");
                    Range->dynamicCall("InsertAfter(QString)", j);
                    Range->dynamicCall("InsertAfter(QString)", " ");
                    Range->dynamicCall("InsertAfter(QString)", filename);

                    Range->dynamicCall("InsertAfter(QString)", "\v\v");
                    Range->dynamicCall("InsertAfter(QString)", text);
                    Range->dynamicCall("InsertAfter(QString)", "\v");
                }
                else
                {
                   j = 1;

                   Range->dynamicCall("InsertAfter(QString)", i);
                   Range->dynamicCall("InsertAfter(QString)", ". ");
                   Range->dynamicCall("InsertAfter(QString)", dirname);
                   Range->dynamicCall("InsertAfter(QString)", "\v");
                   Range->dynamicCall("InsertAfter(QString)", i);
                   Range->dynamicCall("InsertAfter(QString)", ".");
                   Range->dynamicCall("InsertAfter(QString)", j);
                   Range->dynamicCall("InsertAfter(QString)", ". ");
                   Range->dynamicCall("InsertAfter(QString)", filename);

                   Range->dynamicCall("InsertAfter(QString)", "\v\v");
                   Range->dynamicCall("InsertAfter(QString)", text);
                   Range->dynamicCall("InsertAfter(QString)", "\v");

                   i++;
                   previousDirname = dirname;

                }
            }
        }
        WordApplication->setProperty("Visible", true); // Делаем Word видимым
        titleDoc->dynamicCall("SaveAs(QString)", "New.docx");

        delete Range;
        delete ActiveDocument;
        delete WordDocuments;
        delete WordApplication;

   }




