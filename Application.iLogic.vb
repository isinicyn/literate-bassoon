Option Explicit
AddReference "System.Windows.Forms"
AddReference "System.Drawing"
Imports WinForms = System.Windows.Forms
Imports Drawing = System.Drawing
Imports IO = System.IO
Imports Regex = System.Text.RegularExpressions.Regex
Imports ColorTranslator = System.Drawing.ColorTranslator


Sub Main()
    ' Определение местоположения файла сборки
    Dim assemblyFilePath As String = ThisDoc.Document.FullFileName
    Dim assemblyDirectory As String = IO.Path.GetDirectoryName(assemblyFilePath)
    Dim baseDirectory As String = assemblyDirectory

    ' Создание формы
    Dim form As New WinForms.Form()
    form.Text = "Выбор опций"
    form.StartPosition = WinForms.FormStartPosition.CenterScreen
    form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.Size = New Drawing.Size(1000, 800)
    form.AutoSize = True
    form.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink

    Dim defaultFont As New Drawing.Font("Segoe UI", 9)
    form.Font = defaultFont
    form.BackColor = Drawing.Color.WhiteSmoke

    ' Создание TabControl для вкладок
    Dim tabControl As New WinForms.TabControl()
    tabControl.Dock = WinForms.DockStyle.Fill
    form.Controls.Add(tabControl)

    ' Создание вкладки для основного скрипта
    Dim tabMain As New WinForms.TabPage("Основные операции")
    Dim tabAdditional As New WinForms.TabPage("Доп. операции")
    tabMain.BackColor = form.BackColor
    tabAdditional.BackColor = form.BackColor
    tabControl.TabPages.Add(tabMain)
    tabControl.TabPages.Add(tabAdditional)

    ' Основной макет для первой вкладки
    Dim mainLayout As New WinForms.TableLayoutPanel()
    mainLayout.Dock = WinForms.DockStyle.Fill
    mainLayout.RowCount = 4
    mainLayout.RowStyles.Add(New WinForms.RowStyle())
    mainLayout.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 40))
    mainLayout.RowStyles.Add(New WinForms.RowStyle())
    mainLayout.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 60))
    tabMain.Controls.Add(mainLayout)

    ' Вертикальный макет для второй вкладки
    Dim additionalFlow As New WinForms.FlowLayoutPanel()
    additionalFlow.Dock = WinForms.DockStyle.Fill
    additionalFlow.FlowDirection = WinForms.FlowDirection.TopDown
    additionalFlow.WrapContents = False
    additionalFlow.AutoScroll = True
    tabAdditional.Controls.Add(additionalFlow)

    ' Создание GroupBox для отображения пути в основной вкладке
    Dim gbPath As New WinForms.GroupBox()
    gbPath.Text = "Проверяется папка:"
    StyleGroupBox(gbPath)
    mainLayout.Controls.Add(gbPath, 0, 0)

    ' Создание RichTextBox для отображения пути
    Dim rtbPath As New WinForms.RichTextBox()
    rtbPath.Top = 30
    rtbPath.Left = 20
    rtbPath.Width = 840
    rtbPath.Height = 20
    rtbPath.ReadOnly = True
    rtbPath.BorderStyle = WinForms.BorderStyle.None
    rtbPath.ScrollBars = WinForms.RichTextBoxScrollBars.None
    rtbPath.Multiline = False
    gbPath.Controls.Add(rtbPath)

    ' Создание кнопки для выбора папки
    Dim btnBrowse As New WinForms.Button()
    btnBrowse.Text = "Выбрать..."
    btnBrowse.Top = 28
    btnBrowse.Left = 870
    btnBrowse.Width = 80
    btnBrowse.Height = 24
    StyleButton(btnBrowse)
    gbPath.Controls.Add(btnBrowse)

    ' Создание кнопки для возврата изначального пути
    Dim btnAuto As New WinForms.Button()
    btnAuto.Text = "Авто"
    btnAuto.Top = 58
    btnAuto.Left = 870
    btnAuto.Width = 80
    btnAuto.Height = 24
    StyleButton(btnAuto)
    gbPath.Controls.Add(btnAuto)

    ' Создание метки для статуса папки
    Dim lblStatus As New WinForms.Label()
    lblStatus.Top = 60
    lblStatus.Left = 20
    lblStatus.Width = 840
    lblStatus.Height = 20
    lblStatus.ForeColor = ColorTranslator.FromHtml("#ee5c0d")
    gbPath.Controls.Add(lblStatus)

    ' Начальное обновление пути и статуса папки
    UpdatePathAndStatus(baseDirectory, rtbPath, lblStatus)

    ' Создание GroupBox для выбора опций в основной вкладке
    Dim gbOptions As New WinForms.GroupBox()
    gbOptions.Text = "Выбор опций"
    StyleGroupBox(gbOptions)

    ' Создание панели для кнопок выбора/снятия
    Dim panelSelectButtons As New WinForms.FlowLayoutPanel()
    panelSelectButtons.Top = 20
    panelSelectButtons.Left = 10
    panelSelectButtons.Width = 440
    panelSelectButtons.AutoSize = True
    panelSelectButtons.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    gbOptions.Controls.Add(panelSelectButtons)

    ' Создание кнопки для выбора всех чекбоксов
    Dim btnSelectAll As New WinForms.Button()
    btnSelectAll.Text = "Выбрать всё"
    btnSelectAll.AutoSize = True
    StyleButton(btnSelectAll)

    ' Создание кнопки для снятия выделения со всех чекбоксов
    Dim btnDeselectAll As New WinForms.Button()
    btnDeselectAll.Text = "Снять выделение"
    btnDeselectAll.AutoSize = True
    StyleButton(btnDeselectAll)

    ' Добавление кнопок выбора/снятия в панель
    panelSelectButtons.Controls.Add(btnSelectAll)
    panelSelectButtons.Controls.Add(btnDeselectAll)

    ' Создание панели для чекбоксов
    Dim panelOptions As New WinForms.FlowLayoutPanel()
    panelOptions.Top = panelSelectButtons.Bottom + 10
    panelOptions.Left = 10
    panelOptions.Width = 440
    panelOptions.AutoSize = True
    panelOptions.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    panelOptions.FlowDirection = WinForms.FlowDirection.TopDown
    gbOptions.Controls.Add(panelOptions)

    ' Создание чекбоксов для каждой опции
    Dim cb00 As New WinForms.CheckBox()
    cb00.Text = "[00] Чертежи"
    cb00.AutoSize = True

    Dim cb02 As New WinForms.CheckBox()
    cb02.Text = "[02] Лазерный стол"
    cb02.AutoSize = True

    Dim cb06 As New WinForms.CheckBox()
    cb06.Text = "[06] Лазерный труборез"
    cb06.AutoSize = True

    Dim cb08_1 As New WinForms.CheckBox()
    cb08_1.Text = "[08] Нестинг"
    cb08_1.AutoSize = True

    Dim cb08_2 As New WinForms.CheckBox()
    cb08_2.Text = "[08] Присадка"
    cb08_2.AutoSize = True

    ' Добавление чекбоксов в панель
    panelOptions.Controls.Add(cb00)
    panelOptions.Controls.Add(cb02)
    panelOptions.Controls.Add(cb06)
    panelOptions.Controls.Add(cb08_1)
    panelOptions.Controls.Add(cb08_2)

    ' Создание GroupBox для пользовательских папок в основной вкладке
    Dim gbCustomFolders As New WinForms.GroupBox()
    gbCustomFolders.Text = "Пользовательские папки"
    StyleGroupBox(gbCustomFolders)

    ' Создание панели для пользовательских папок
    Dim panelCustomFolders As New WinForms.FlowLayoutPanel()
    panelCustomFolders.Top = 20
    panelCustomFolders.Left = 10
    panelCustomFolders.Width = 440
    panelCustomFolders.AutoSize = True
    panelCustomFolders.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    panelCustomFolders.FlowDirection = WinForms.FlowDirection.TopDown
    gbCustomFolders.Controls.Add(panelCustomFolders)

    ' Добавление шести полей ввода для пользовательских папок
    For i As Integer = 1 To 6
        Dim tbCustomFolder As New WinForms.TextBox()
        tbCustomFolder.Width = 400
        tbCustomFolder.Enabled = (i = 1)
        AddHandler tbCustomFolder.TextChanged, Sub(tbSender As Object, tbArgs As EventArgs)
                                                   If Not String.IsNullOrWhiteSpace(tbCustomFolder.Text) Then
                                                       Dim nextIndex As Integer = panelCustomFolders.Controls.GetChildIndex(tbSender) + 1
                                                       If nextIndex < panelCustomFolders.Controls.Count Then
                                                           panelCustomFolders.Controls(nextIndex).Enabled = True
                                                       End If
                                                   Else
                                                       Dim thisIndex As Integer = panelCustomFolders.Controls.GetChildIndex(tbSender)
                                                       For j As Integer = thisIndex + 1 To panelCustomFolders.Controls.Count - 1
                                                           panelCustomFolders.Controls(j).Enabled = False
                                                           CType(panelCustomFolders.Controls(j), WinForms.TextBox).Text = ""
                                                       Next
                                                   End If
                                               End Sub
        panelCustomFolders.Controls.Add(tbCustomFolder)
    Next

    ' Контейнер с опциями и пользовательскими папками
    Dim optionsSplit As New WinForms.SplitContainer()
    optionsSplit.Dock = WinForms.DockStyle.Fill
    optionsSplit.FixedPanel = WinForms.FixedPanel.Panel1
    optionsSplit.SplitterDistance = 480
    optionsSplit.Panel1.Controls.Add(gbOptions)
    optionsSplit.Panel2.Controls.Add(gbCustomFolders)
    mainLayout.Controls.Add(optionsSplit, 0, 1)

    ' Создание панели для кнопок
    Dim panelButtons As New WinForms.FlowLayoutPanel()
    panelButtons.AutoSize = True
    panelButtons.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    panelButtons.Dock = WinForms.DockStyle.Fill
    panelButtons.Margin = New WinForms.Padding(5)
    mainLayout.Controls.Add(panelButtons, 0, 2)

    ' Создание кнопки для отправки формы
    Dim btnSubmit As New WinForms.Button()
    btnSubmit.Text = "Создать"
    btnSubmit.AutoSize = True
    StyleButton(btnSubmit)

    ' Создание кнопки для закрытия формы
    Dim btnClose As New WinForms.Button()
    btnClose.Text = "Закрыть"
    btnClose.AutoSize = True
    StyleButton(btnClose)

    ' Создание кнопки для открытия папки
    Dim btnOpenFolder As New WinForms.Button()
    btnOpenFolder.Text = "Открыть папку"
    btnOpenFolder.AutoSize = True
    StyleButton(btnOpenFolder)

    ' Создание кнопки для удаления папки
    Dim btnDeleteAll As New WinForms.Button()
    btnDeleteAll.Text = "Стереть всё"
    btnDeleteAll.AutoSize = True
    btnDeleteAll.Anchor = WinForms.AnchorStyles.Right
    StyleButton(btnDeleteAll)

    ' Добавление кнопок в панель
    panelButtons.Controls.Add(btnSubmit)
    panelButtons.Controls.Add(btnClose)
    panelButtons.Controls.Add(btnOpenFolder)
    panelButtons.Controls.Add(btnDeleteAll)

    ' Создание основной панели для списков
    Dim panelLists As New WinForms.Panel()
    panelLists.AutoSize = True
    panelLists.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    panelLists.Dock = WinForms.DockStyle.Fill
    panelLists.Margin = New WinForms.Padding(5)
    mainLayout.Controls.Add(panelLists, 0, 3)

    ' Создание GroupBox для существующих папок
    Dim gbExistingDirs As New WinForms.GroupBox()
    gbExistingDirs.Text = "Существующие папки"
    gbExistingDirs.ForeColor = Drawing.Color.Black
    gbExistingDirs.Width = 460
    StyleGroupBox(gbExistingDirs)

    ' Создание ListView для отображения существующих папок
    Dim lvExistingDirs As New WinForms.ListView()
    lvExistingDirs.Top = 20
    lvExistingDirs.Left = 10
    lvExistingDirs.Width = 440
    lvExistingDirs.Height = 240
    lvExistingDirs.View = WinForms.View.Details
    lvExistingDirs.Columns.Add("№", 40, WinForms.HorizontalAlignment.Left)
    lvExistingDirs.Columns.Add("Папка", 300, WinForms.HorizontalAlignment.Left)
    lvExistingDirs.Columns.Add("Статус", 100, WinForms.HorizontalAlignment.Left)
    lvExistingDirs.FullRowSelect = True
    lvExistingDirs.Scrollable = True
    lvExistingDirs.MultiSelect = False
    lvExistingDirs.HideSelection = False
    lvExistingDirs.HeaderStyle = WinForms.ColumnHeaderStyle.Clickable
    gbExistingDirs.Controls.Add(lvExistingDirs)

    ' Создание GroupBox для несвязанных папок
    Dim gbUnrelatedDirs As New WinForms.GroupBox()
    gbUnrelatedDirs.Text = "Несвязанные папки"
    gbUnrelatedDirs.ForeColor = Drawing.Color.Black
    gbUnrelatedDirs.Width = 460
    StyleGroupBox(gbUnrelatedDirs)

    ' Создание ListView для отображения несвязанных папок
    Dim lvUnrelatedDirs As New WinForms.ListView()
    lvUnrelatedDirs.Top = 20
    lvUnrelatedDirs.Left = 10
    lvUnrelatedDirs.Width = 440
    lvUnrelatedDirs.Height = 240
    lvUnrelatedDirs.View = WinForms.View.Details
    lvUnrelatedDirs.Columns.Add("№", 40, WinForms.HorizontalAlignment.Left)
    lvUnrelatedDirs.Columns.Add("Папка", 360, WinForms.HorizontalAlignment.Left)
    lvUnrelatedDirs.FullRowSelect = True
    lvUnrelatedDirs.Scrollable = True
    lvUnrelatedDirs.MultiSelect = False
    lvUnrelatedDirs.HideSelection = False
    lvUnrelatedDirs.HeaderStyle = WinForms.ColumnHeaderStyle.Clickable
    gbUnrelatedDirs.Controls.Add(lvUnrelatedDirs)

    ' Контейнер для списков директорий
    Dim listsSplit As New WinForms.SplitContainer()
    listsSplit.Dock = WinForms.DockStyle.Fill
    listsSplit.FixedPanel = WinForms.FixedPanel.Panel1
    listsSplit.SplitterDistance = 480
    listsSplit.Panel1.Controls.Add(gbExistingDirs)
    listsSplit.Panel2.Controls.Add(gbUnrelatedDirs)
    panelLists.Controls.Add(listsSplit)

    ' Изначально скрываем панели, если "- Заявки" не существует
    If Not IO.Directory.Exists(IO.Path.Combine(baseDirectory, "- Заявки")) Then
        gbExistingDirs.Visible = False
        gbUnrelatedDirs.Visible = False
        panelLists.Visible = False
        btnOpenFolder.Enabled = False
        btnDeleteAll.Enabled = False
    Else
        UpdateDirectoryLists(baseDirectory, lvExistingDirs, lvUnrelatedDirs, cb00, cb02, cb06, cb08_1, cb08_2)
        gbExistingDirs.Visible = True
        gbUnrelatedDirs.Visible = True
        panelLists.Visible = True
        btnOpenFolder.Enabled = True
        btnDeleteAll.Enabled = True
    End If

    ' Обработчики событий для основной вкладки
    AddHandler btnBrowse.Click, Sub(sender As Object, e As EventArgs)
                                    Dim folderDialog As New WinForms.FolderBrowserDialog()
                                    folderDialog.Description = "Выберите папку для создания субдиректорий"
                                    folderDialog.SelectedPath = baseDirectory
                                    If folderDialog.ShowDialog() = WinForms.DialogResult.OK Then
                                        baseDirectory = folderDialog.SelectedPath
                                        UpdatePathAndStatus(baseDirectory, rtbPath, lblStatus)
                                        ' Проверяем наличие субдиректории "- Заявки"
                                        If IO.Directory.Exists(IO.Path.Combine(baseDirectory, "- Заявки")) Then
                                            ' Обновляем списки после выбора новой папки
                                            UpdateDirectoryLists(baseDirectory, lvExistingDirs, lvUnrelatedDirs, cb00, cb02, cb06, cb08_1, cb08_2)
                                            gbExistingDirs.Visible = True
                                            gbUnrelatedDirs.Visible = True
                                            panelLists.Visible = True
                                            btnOpenFolder.Enabled = True
                                            btnDeleteAll.Enabled = True
                                        Else
                                            ' Скрыть панели с директориями, если "- Заявки" не существует
                                            gbExistingDirs.Visible = False
                                            gbUnrelatedDirs.Visible = False
                                            panelLists.Visible = False
                                            btnOpenFolder.Enabled = False
                                            btnDeleteAll.Enabled = False
                                        End If
                                    End If
                                End Sub
    AddHandler btnAuto.Click, Sub(sender As Object, e As EventArgs)
                                  baseDirectory = assemblyDirectory
                                  UpdatePathAndStatus(baseDirectory, rtbPath, lblStatus)
                                  ' Проверяем наличие субдиректории "- Заявки"
                                  If IO.Directory.Exists(IO.Path.Combine(baseDirectory, "- Заявки")) Then
                                      ' Обновляем списки после возврата изначального пути
                                      UpdateDirectoryLists(baseDirectory, lvExistingDirs, lvUnrelatedDirs, cb00, cb02, cb06, cb08_1, cb08_2)
                                      gbExistingDirs.Visible = True
                                      gbUnrelatedDirs.Visible = True
                                      panelLists.Visible = True
                                      btnOpenFolder.Enabled = True
                                      btnDeleteAll.Enabled = True
                                  Else
                                      ' Скрыть панели с директориями, если "- Заявки" не существует
                                      gbExistingDirs.Visible = False
                                      gbUnrelatedDirs.Visible = False
                                      panelLists.Visible = False
                                      btnOpenFolder.Enabled = False
                                      btnDeleteAll.Enabled = False
                                  End If
                              End Sub
    AddHandler btnSubmit.Click, Sub(sender As Object, e As EventArgs)
                                    If Not (cb00.Checked Or cb02.Checked Or cb06.Checked Or cb08_1.Checked Or cb08_2.Checked Or CustomFolderExists(panelCustomFolders)) Then
                                        WinForms.MessageBox.Show("Пожалуйста, выберите хотя бы одну опцию для создания папок.", "Ошибка", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                                        Return
                                    End If
                                    HandleButtonClick(sender, cb00, cb02, cb06, cb08_1, cb08_2, baseDirectory, lvExistingDirs, lvUnrelatedDirs, panelCustomFolders, lblStatus, form, gbPath, gbOptions, gbCustomFolders, panelButtons, panelLists, btnOpenFolder, btnDeleteAll)
                                End Sub
    AddHandler btnSelectAll.Click, Sub(sender As Object, e As EventArgs)
                                       cb00.Checked = True
                                       cb02.Checked = True
                                       cb06.Checked = True
                                       cb08_1.Checked = True
                                       cb08_2.Checked = True
                                   End Sub
    AddHandler btnDeselectAll.Click, Sub(sender As Object, e As EventArgs)
                                         cb00.Checked = False
                                         cb02.Checked = False
                                         cb06.Checked = False
                                         cb08_1.Checked = False
                                         cb08_2.Checked = False
                                     End Sub
    AddHandler btnClose.Click, Sub(sender As Object, e As EventArgs)
                                   form.Close()
                               End Sub
    AddHandler btnOpenFolder.Click, Sub(sender As Object, e As EventArgs)
                                        Dim targetPath As String = IO.Path.Combine(baseDirectory, "- Заявки")
                                        If IO.Directory.Exists(targetPath) Then
                                            Process.Start("explorer.exe", targetPath)
                                        Else
                                            WinForms.MessageBox.Show("Папка назначения не найдена.", "Ошибка", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                                        End If
                                    End Sub

    AddHandler btnDeleteAll.Click, Sub(sender As Object, e As EventArgs)
                                       Dim result As WinForms.DialogResult = WinForms.MessageBox.Show("Папка '- Заявки' и всё её содержимое будет удалено без возможности восстановления. Вы уверены?", "Подтверждение удаления", WinForms.MessageBoxButtons.OKCancel, WinForms.MessageBoxIcon.Warning)
                                       If result = WinForms.DialogResult.OK Then
                                           DeleteFolder(IO.Path.Combine(baseDirectory, "- Заявки"))
                                           UpdatePathAndStatus(baseDirectory, rtbPath, lblStatus)
                                           gbExistingDirs.Visible = False
                                           gbUnrelatedDirs.Visible = False
                                           panelLists.Visible = False
                                           btnOpenFolder.Enabled = False
                                           btnDeleteAll.Enabled = False
                                       End If
                                   End Sub

    ' --- Реализация второй вкладки ---
    ' Создание GroupBox для операций с PDF
    Dim gbPDFOperations As New WinForms.GroupBox()
    gbPDFOperations.Text = "Операции с PDF"
    gbPDFOperations.Dock = WinForms.DockStyle.Top
    StyleGroupBox(gbPDFOperations)
    additionalFlow.Controls.Add(gbPDFOperations)

    ' Создание кнопки для сканирования и перемещения PDF файлов
    Dim btnScanPDF As New WinForms.Button()
    btnScanPDF.Text = "Сканировать PDF"
    btnScanPDF.Top = 20
    btnScanPDF.Left = 20
    btnScanPDF.Width = 200
    btnScanPDF.Height = 30
    StyleButton(btnScanPDF)
    gbPDFOperations.Controls.Add(btnScanPDF)

    Dim btnMovePDF As New WinForms.Button()
    btnMovePDF.Text = "Переместить PDF"
    btnMovePDF.Top = 20
    btnMovePDF.Left = 240
    btnMovePDF.Width = 200
    btnMovePDF.Height = 30
    StyleButton(btnMovePDF)
    gbPDFOperations.Controls.Add(btnMovePDF)

    ' Создание метки для отображения статистики
    Dim lblStats As New WinForms.Label()
    lblStats.Top = 60
    lblStats.Left = 20
    lblStats.Width = 400
    lblStats.Height = 300
    gbPDFOperations.Controls.Add(lblStats)

    ' Обработчик события для кнопки сканирования PDF файлов
    AddHandler btnScanPDF.Click, Sub(sender As Object, e As EventArgs)
                                     Dim stats As String = ScanPDFs(baseDirectory)
                                     lblStats.Text = stats
                                 End Sub

    ' Обработчик события для кнопки перемещения PDF файлов
    AddHandler btnMovePDF.Click, Sub(sender As Object, e As EventArgs)
                                     Dim baseDir As String = IO.Path.Combine(baseDirectory, "- Заявки")
                                     Dim currentDateDir As String = IO.Path.Combine(baseDir, DateTime.Now.ToString("dd.MM.yyyy"))
                                     If Not IO.Directory.Exists(currentDateDir) Then
                                         Dim result As WinForms.DialogResult = WinForms.MessageBox.Show("Папка заявок на текущую дату не существует. Перейдите на вкладку 'Основные операции', чтобы создать её.", "Предупреждение", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                                         If result = WinForms.DialogResult.OK Then
                                             tabControl.SelectedTab = tabMain
                                             Return
                                         Else
                                             Return
                                         End If
                                     End If
                                     Dim stats As String = MovePDFs(baseDirectory)
                                     lblStats.Text = stats
                                 End Sub

    ' Создание GroupBox для проверки наполнения папок
    Dim gbFolderCheck As New WinForms.GroupBox()
    gbFolderCheck.Text = "Проверка наполнения папок"
    gbFolderCheck.Dock = WinForms.DockStyle.Top
    StyleGroupBox(gbFolderCheck)
    additionalFlow.Controls.Add(gbFolderCheck)

    ' Создание кнопки для проверки папок
    Dim btnCheckFolders As New WinForms.Button()
    btnCheckFolders.Text = "Проверить папки"
    btnCheckFolders.Top = 20
    btnCheckFolders.Left = 20
    btnCheckFolders.Width = 200
    btnCheckFolders.Height = 30
    StyleButton(btnCheckFolders)
    gbFolderCheck.Controls.Add(btnCheckFolders)

    ' Создание метки для результата проверки
    Dim lblCheckResult As New WinForms.Label()
    lblCheckResult.Top = 60
    lblCheckResult.Left = 20
    lblCheckResult.Width = 400
    lblCheckResult.Height = 30
    gbFolderCheck.Controls.Add(lblCheckResult)

    ' Обработчик события для кнопки проверки папок
    AddHandler btnCheckFolders.Click, Sub(sender As Object, e As EventArgs)
                                          Dim baseDir As String = IO.Path.Combine(baseDirectory, "- Заявки")
                                          Dim currentDateDir As String = IO.Path.Combine(baseDir, DateTime.Now.ToString("dd.MM.yyyy"))
                                          If Not IO.Directory.Exists(currentDateDir) Then
                                              lblCheckResult.Text = "Папка заявок на текущую дату отсутствует."
                                              lblCheckResult.ForeColor = ColorTranslator.FromHtml("#d90429")
                                              Return
                                          End If

                                          Dim emptyDirs As List(Of String) = CheckEmptyFolders(currentDateDir)
                                          If emptyDirs.Count > 0 Then
                                              Dim emptyDirsForm As New WinForms.Form()
                                              emptyDirsForm.Text = "Пустые папки"
                                              emptyDirsForm.StartPosition = WinForms.FormStartPosition.CenterParent
                                              emptyDirsForm.AutoSize = True
                                              emptyDirsForm.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink

                                              Dim lvEmptyDirs As New WinForms.ListView()
                                              lvEmptyDirs.Top = 20
                                              lvEmptyDirs.Left = 10
                                              lvEmptyDirs.Width = 360
                                              lvEmptyDirs.Height = 200
                                              lvEmptyDirs.View = WinForms.View.Details
                                              lvEmptyDirs.Columns.Add("№", 40, WinForms.HorizontalAlignment.Left)
                                              lvEmptyDirs.Columns.Add("Папка", 300, WinForms.HorizontalAlignment.Left)
                                              lvEmptyDirs.FullRowSelect = True
                                              lvEmptyDirs.Scrollable = True
                                              lvEmptyDirs.MultiSelect = False
                                              lvEmptyDirs.HideSelection = False
                                              lvEmptyDirs.HeaderStyle = WinForms.ColumnHeaderStyle.Clickable

                                              Dim count As Integer = 1
                                              For Each dir As String In emptyDirs
                                                  Dim item As New WinForms.ListViewItem(count.ToString())
                                                  item.SubItems.Add(Dir)
                                                  lvEmptyDirs.Items.Add(item)
                                                  count += 1
                                              Next

                                              emptyDirsForm.Controls.Add(lvEmptyDirs)
                                              emptyDirsForm.ShowDialog()
                                          Else
                                              lblCheckResult.Text = "Не найдено!"
                                              lblCheckResult.ForeColor = ColorTranslator.FromHtml("#02bd11")
                                          End If
                                      End Sub

    ' Показ формы
    WinForms.Application.Run(form)
End Sub

' Функция для обновления пути и статуса папки
Sub UpdatePathAndStatus(baseDirectory As String, rtbPath As WinForms.RichTextBox, lblStatus As WinForms.Label)
    Dim fullPath As String = baseDirectory & "\- Заявки"
    rtbPath.Text = fullPath

    Dim boldStart As Integer = rtbPath.Text.IndexOf("- Заявки")
    Dim boldLength As Integer = "- Заявки".Length

    rtbPath.Select(boldStart, boldLength)
    rtbPath.SelectionFont = New Drawing.Font(rtbPath.Font, Drawing.FontStyle.Bold)
    rtbPath.Select(0, 0) ' Сбросить выделение

    If IO.Directory.Exists(IO.Path.Combine(baseDirectory, "- Заявки")) Then
        lblStatus.Text = "Папка обнаружена"
    Else
        lblStatus.Text = "Папка заявок отсутствует. Будет создана автоматически"
    End If
End Sub

' Функция для обновления списков директорий
Sub UpdateDirectoryLists(baseDirectory As String, lvExistingDirs As WinForms.ListView, lvUnrelatedDirs As WinForms.ListView, cb00 As WinForms.CheckBox, cb02 As WinForms.CheckBox, cb06 As WinForms.CheckBox, cb08_1 As WinForms.CheckBox, cb08_2 As WinForms.CheckBox)
    lvExistingDirs.Items.Clear()
    lvUnrelatedDirs.Items.Clear()
    
    Dim baseDir As String = IO.Path.Combine(baseDirectory, "- Заявки")
    Dim existingDirs = IO.Directory.GetDirectories(baseDir, "*", IO.SearchOption.TopDirectoryOnly)
    If existingDirs.Any() Then
        Dim dirInfos = existingDirs.Select(Function(d) New IO.DirectoryInfo(d)).OrderBy(Function(d) d.Name).ToList()
        Dim currentDateDirName As String = DateTime.Now.ToString("dd.MM.yyyy")
        Dim regex As New Regex("^\d{2}\.\d{2}\.\d{4}$")

        Dim count As Integer = 1
        For Each dirInfo In dirInfos
            Dim item As New WinForms.ListViewItem(count.ToString())
            item.SubItems.Add("⯈ " & dirInfo.Name)
            If regex.IsMatch(dirInfo.Name) Then
                If dirInfo.Name = currentDateDirName Then
                    item.SubItems.Add("- Данные на текущую дату уже есть!")
                    item.UseItemStyleForSubItems = False
                    item.SubItems(2).ForeColor = ColorTranslator.FromHtml("#ef233c")
                    item.SubItems(2).Font = New Drawing.Font(item.Font, Drawing.FontStyle.Bold)
                Else
                    item.SubItems.Add("")
                End If
                lvExistingDirs.Items.Add(item)
            Else
                lvUnrelatedDirs.Items.Add(item)
            End If
            count += 1
        Next

        ' Автоматическое включение чекбоксов, если подпапки существуют
        cb00.Checked = IO.Directory.Exists(IO.Path.Combine(baseDir, currentDateDirName, "[00] Чертежи"))
        cb02.Checked = IO.Directory.Exists(IO.Path.Combine(baseDir, currentDateDirName, "[02] Лазерный стол"))
        cb06.Checked = IO.Directory.Exists(IO.Path.Combine(baseDir, currentDateDirName, "[06] Лазерный труборез"))
        cb08_1.Checked = IO.Directory.Exists(IO.Path.Combine(baseDir, currentDateDirName, "[08] Нестинг"))
        cb08_2.Checked = IO.Directory.Exists(IO.Path.Combine(baseDir, currentDateDirName, "[08] Присадка"))

        lvExistingDirs.AutoResizeColumns(WinForms.ColumnHeaderAutoResizeStyle.ColumnContent)
        lvExistingDirs.AutoResizeColumns(WinForms.ColumnHeaderAutoResizeStyle.HeaderSize)
        lvUnrelatedDirs.AutoResizeColumns(WinForms.ColumnHeaderAutoResizeStyle.ColumnContent)
        lvUnrelatedDirs.AutoResizeColumns(WinForms.ColumnHeaderAutoResizeStyle.HeaderSize)
    End If
End Sub

' Функция для создания папок
Sub CreateFolder(folderPath As String)
    If Not IO.Directory.Exists(folderPath) Then
        IO.Directory.CreateDirectory(folderPath)
    End If
End Sub

' Функция для удаления папок
Sub DeleteFolder(folderPath As String)
    If IO.Directory.Exists(folderPath) Then
        IO.Directory.Delete(folderPath, True)
    End If
End Sub

' Функция для проверки наличия пользовательской папки
Function CustomFolderExists(panelCustomFolders As WinForms.FlowLayoutPanel) As Boolean
    For Each control As WinForms.Control In panelCustomFolders.Controls
        If TypeOf Control Is WinForms.TextBox AndAlso Not String.IsNullOrWhiteSpace(Control.Text) Then
            Return True
        End If
    Next
    Return False
End Function

' Общий обработчик для кнопки "Создать"
Sub HandleButtonClick(sender As Object, cb00 As WinForms.CheckBox, cb02 As WinForms.CheckBox, cb06 As WinForms.CheckBox, cb08_1 As WinForms.CheckBox, cb08_2 As WinForms.CheckBox, baseDirectory As String, lvExistingDirs As WinForms.ListView, lvUnrelatedDirs As WinForms.ListView, panelCustomFolders As WinForms.FlowLayoutPanel, lblStatus As WinForms.Label, form As WinForms.Form, gbPath As WinForms.GroupBox, gbOptions As WinForms.GroupBox, gbCustomFolders As WinForms.GroupBox, panelButtons As WinForms.FlowLayoutPanel, panelLists As WinForms.Panel, btnOpenFolder As WinForms.Button, btnDeleteAll As WinForms.Button)
    Dim baseDir As String = IO.Path.Combine(baseDirectory, "- Заявки")

    ' Создание субдиректории с текущей датой
    Dim currentDateDir As String = IO.Path.Combine(baseDir, DateTime.Now.ToString("dd.MM.yyyy"))

    If IO.Directory.Exists(currentDateDir) Then
        ' Создание формы предупреждения
        Dim warningForm As New WinForms.Form()
        warningForm.Text = "Предупреждение"
        warningForm.StartPosition = WinForms.FormStartPosition.CenterParent
        warningForm.AutoSize = True
        warningForm.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink

        ' Создание GroupBox для опций выбора действия
        Dim gbWarning As New WinForms.GroupBox()
        gbWarning.Text = "На текущую дату уже есть данные. Выберите действие:"
        StyleGroupBox(gbWarning)
        gbWarning.Width = 340
        gbWarning.AutoSize = False
        warningForm.Controls.Add(gbWarning)

        Dim cbDeleteAll As New WinForms.CheckBox()
        cbDeleteAll.Text = "Удалить всё и создать заново"
        cbDeleteAll.Top = 30
        cbDeleteAll.Left = 10
        cbDeleteAll.AutoSize = True

        Dim cbAppend As New WinForms.CheckBox()
        cbAppend.Text = "Дозаписать недостающие директории"
        cbAppend.Top = 60
        cbAppend.Left = 10
        cbAppend.AutoSize = True

        ' Событие для обеспечения выбора только одного чекбокса
        AddHandler cbDeleteAll.CheckedChanged, Sub(s, args)
                                                   If cbDeleteAll.Checked Then cbAppend.Checked = False
                                               End Sub

        AddHandler cbAppend.CheckedChanged, Sub(s, args)
                                                If cbAppend.Checked Then cbDeleteAll.Checked = False
                                            End Sub

        gbWarning.Controls.Add(cbDeleteAll)
        gbWarning.Controls.Add(cbAppend)

        ' Создание кнопок подтверждения и отмены
        Dim btnConfirm As New WinForms.Button()
        btnConfirm.Text = "Подтвердить"
        btnConfirm.Top = gbWarning.Bottom + 10
        btnConfirm.Left = 10
        btnConfirm.AutoSize = True

        Dim btnCancel As New WinForms.Button()
        btnCancel.Text = "Отмена"
        btnCancel.Top = gbWarning.Bottom + 10
        btnCancel.Left = gbWarning.Width - btnCancel.Width - 20
        btnCancel.AutoSize = True

        AddHandler btnConfirm.Click, Sub(s, args)
                                         If cbDeleteAll.Checked Then
                                             ' Удалить всё и создать заново
                                             DeleteFolder(currentDateDir)
                                             CreateFolder(currentDateDir)
                                         ElseIf cbAppend.Checked Then
                                             ' Дозаписать недостающие директории
                                             ' Только создать недостающие папки, так что оставляем директорию как есть
                                         Else
                                             ' Если ничего не выбрано, просто закрыть форму предупреждения
                                             warningForm.Close()
                                             Return
                                         End If
                                         warningForm.Close()
                                     End Sub

        AddHandler btnCancel.Click, Sub(s, args)
                                        warningForm.Close()
                                    End Sub

        warningForm.Controls.Add(btnConfirm)
        warningForm.Controls.Add(btnCancel)

        warningForm.ShowDialog()
    Else
        ' Создание новой субдиректории
        CreateFolder(currentDateDir)
        ' Обновление статуса после успешного создания папок
        lblStatus.Text = "Операция выполнена успешно"
        lblStatus.ForeColor = ColorTranslator.FromHtml("#02bd11")
    End If

    If cb00.Checked Then CreateFolder(IO.Path.Combine(currentDateDir, "[00] Чертежи"))
    If cb02.Checked Then CreateFolder(IO.Path.Combine(currentDateDir, "[02] Лазерный стол"))
    If cb06.Checked Then CreateFolder(IO.Path.Combine(currentDateDir, "[06] Лазерный труборез"))
    If cb08_1.Checked Then CreateFolder(IO.Path.Combine(currentDateDir, "[08] Нестинг"))
    If cb08_2.Checked Then CreateFolder(IO.Path.Combine(currentDateDir, "[08] Присадка"))

    ' Создание пользовательских папок
    For Each control As WinForms.Control In panelCustomFolders.Controls
        If TypeOf Control Is WinForms.TextBox AndAlso Not String.IsNullOrWhiteSpace(Control.Text) Then
            CreateFolder(IO.Path.Combine(currentDateDir, Control.Text))
        End If
    Next

    ' Обновляем списки после создания папок
    UpdateDirectoryLists(baseDirectory, lvExistingDirs, lvUnrelatedDirs, cb00, cb02, cb06, cb08_1, cb08_2)

    ' Обновление состояния кнопок "Открыть папку" и "Стереть всё"
    If IO.Directory.Exists(IO.Path.Combine(baseDirectory, "- Заявки")) Then
        btnOpenFolder.Enabled = True
        btnDeleteAll.Enabled = True
    Else
        btnOpenFolder.Enabled = False
        btnDeleteAll.Enabled = False
    End If
End Sub

' Функция для сканирования PDF файлов
Function ScanPDFs(baseDirectory As String) As String
    Dim pdfFiles = IO.Directory.GetFiles(baseDirectory, "*.pdf", IO.SearchOption.AllDirectories).Where(Function(f) Not f.Contains("\- Заявки\")).ToArray()
    Dim duplicateCount As Integer = 0
    Dim fileCount As Integer = pdfFiles.Length

    Dim fileNames As New HashSet(Of String)()
    For Each pdfFile In pdfFiles
        Dim fileName As String = IO.Path.GetFileName(pdfFile)
        If fileNames.Contains(fileName) Then
            duplicateCount += 1
        Else
            fileNames.Add(fileName)
        End If
    Next

    Dim stats As String = String.Format("Найдено файлов: {0}" & vbCrLf & "Файлов с дубликатами: {1}", fileCount, duplicateCount)
    Return stats
End Function

' Функция для перемещения PDF файлов
Function MovePDFs(baseDirectory As String) As String
    Dim baseDir As String = IO.Path.Combine(baseDirectory, "- Заявки")
    Dim currentDateDir As String = IO.Path.Combine(baseDir, DateTime.Now.ToString("dd.MM.yyyy"))
    Dim targetDir As String = IO.Path.Combine(currentDateDir, "[00] Чертежи")
    CreateFolder(targetDir)

    Dim pdfFiles = IO.Directory.GetFiles(baseDirectory, "*.pdf", IO.SearchOption.AllDirectories).Where(Function(f) Not f.Contains("\- Заявки\")).ToArray()
    Dim movedFiles As New List(Of String)
    Dim duplicateCount As Integer = 0
    Dim fileCount As Integer = 0

    For Each pdfFile In pdfFiles
        Dim fileName As String = IO.Path.GetFileNameWithoutExtension(pdfFile)
        Dim fileExtension As String = IO.Path.GetExtension(pdfFile)
        Dim destFile As String = IO.Path.Combine(targetDir, fileName & fileExtension)
        Dim suffix As Integer = 1
        While IO.File.Exists(destFile)
            duplicateCount += 1
            destFile = IO.Path.Combine(targetDir, fileName & "_(" & suffix.ToString("00") & ")" & fileExtension)
            suffix += 1
        End While
        IO.File.Move(pdfFile, destFile)
        fileCount += 1
    Next

    Dim stats As String = String.Format("Перемещено файлов: {0}" & vbCrLf & "Файлов с дубликатами: {1}", fileCount, duplicateCount)
    Return stats
End Function

' Функция для проверки пустых папок
Function CheckEmptyFolders(currentDateDir As String) As List(Of String)
    Dim emptyDirs As New List(Of String)
    Dim directories = IO.Directory.GetDirectories(currentDateDir)
    For Each dir As String In directories
        If Not IO.Directory.EnumerateFileSystemEntries(Dir).Any() Then
            emptyDirs.Add(IO.Path.GetFileName(Dir))
        End If
    Next
    Return emptyDirs
End Function

' Применение упрощённого стиля к кнопкам
Sub StyleButton(btn As WinForms.Button)
    btn.FlatStyle = WinForms.FlatStyle.Flat
    btn.FlatAppearance.BorderSize = 0
    btn.BackColor = ColorTranslator.FromHtml("#007bff")
    btn.ForeColor = Drawing.Color.White
    btn.Padding = New WinForms.Padding(6, 4, 6, 4)
End Sub

' Применение единого стиля к GroupBox
Sub StyleGroupBox(gb As WinForms.GroupBox)
    gb.AutoSize = True
    gb.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    gb.Dock = WinForms.DockStyle.Fill
    gb.Margin = New WinForms.Padding(5)
End Sub
