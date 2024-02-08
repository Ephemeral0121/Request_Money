import sys
import os
from PyQt5.QtWidgets import QMessageBox, QScrollArea, QSpinBox, QDateEdit, QApplication, QHBoxLayout, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QComboBox, QRadioButton, QButtonGroup, QGroupBox, QFormLayout
from openpyxl import load_workbook
from datetime import datetime
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QPalette, QIcon

class Request_Money(QWidget):

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout()
        mainLayout.setSpacing(10)  # 위젯 간의 간격 조정
        mainLayout.setContentsMargins(10, 10, 10, 10)  # 레이아웃의 여백 조정
        self.setWindowIcon(QIcon(self.resource_path("icon.png")))

        # 엔터티 유형 및 계산 유형 설정
        self.setupEntityTypeGroup(mainLayout)
        self.setupCalculationTypeGroup(mainLayout)

        # 수입금액 입력
        incomeLayout = QHBoxLayout()
        incomeLayout.addWidget(QLabel('기준금액:'))
        self.incomeEdit = QLineEdit()
        incomeLayout.addWidget(self.incomeEdit)
        mainLayout.addLayout(incomeLayout)

        # 원가계산 누진 및 기본계산 누진 설정
        self.setupCostProgression(mainLayout)
        self.setupBaseCalcProgression(mainLayout)

        # 인원수 설정
        self.setupNumPeople(mainLayout)

        # 결과 라벨 구성
        self.resultLabel = QLabel('기본보수:')
        self.resultLabel.setStyleSheet("font-weight: bold; color: green")

        self.additionalLabel = QLabel('')
        self.additionalLabel.setStyleSheet("font-weight: bold; color: green")

        self.final_resultLabel = QLabel('최종기본보수:')
        self.final_resultLabel.setStyleSheet("font-weight: bold; color: green")

        self.costProgressionLabel = QLabel('\n원가계산누진:')
        self.costProgressionLabel.setStyleSheet("font-weight: bold; color: blue")


        # 결과 라벨을 담을 스크롤 영역 설정
        self.scrollArea = QScrollArea()
        self.scrollArea.setWidgetResizable(True)
        self.scrollAreaWidgetContents = QWidget()
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.resultLayout = QVBoxLayout(self.scrollAreaWidgetContents)
        self.scrollAreaWidgetContents.setLayout(self.resultLayout)

        # 기본보수, 추가보수, 최종기본보수 라벨을 스크롤 영역에 추가
        self.resultLayout.addWidget(self.resultLabel)
        self.resultLayout.addWidget(self.additionalLabel)
        self.resultLayout.addWidget(self.final_resultLabel)
        self.resultLayout.addWidget(self.costProgressionLabel)

        # 스크롤 영역을 메인 레이아웃에 추가
        mainLayout.addWidget(self.scrollArea)

        # 회사명 입력
        companyNameLayout = QHBoxLayout()
        companyNameLayout.addWidget(QLabel('회사명:'))
        self.companyNameEdit = QLineEdit()
        companyNameLayout.addWidget(self.companyNameEdit)
        mainLayout.addLayout(companyNameLayout)

        # 납기일 설정
        self.setupDueDate(mainLayout)

        # 엑셀 파일 생성 버튼
        self.createExcelButton = QPushButton('엑셀 파일 생성')
        mainLayout.addWidget(self.createExcelButton)

        # 모든 위젯 설정 후 이벤트 연결
        self.incomeEdit.textChanged.connect(self.calculateRemuneration)  # 입력 값 변경 시 계산 메소드 호출
        self.corporateButton.toggled.connect(self.calculateRemuneration)
        self.individualButton.toggled.connect(self.calculateRemuneration)
        self.typeAButton.toggled.connect(self.calculateRemuneration)
        self.typeBButton.toggled.connect(self.calculateRemuneration)
        self.createExcelButton.clicked.connect(self.create_excel_file)

        self.corporateButton.setChecked(True)
        self.typeBButton.setChecked(True)
        self.costProgressionGroup.buttons()[0].setChecked(True)
        self.baseCalcProgressionGroup.buttons()[0].setChecked(True)

        # 메시지 라벨
        self.messageLabel = QLabel()
        mainLayout.addWidget(self.messageLabel)

        self.applyStyleSheet()
        self.setGeometry(300, 300, 600, 1000)
        self.setLayout(mainLayout)
        self.setWindowTitle('조정보수청구_2023')

    def applyStyleSheet(self):
        self.setStyleSheet("""
            QWidget {
                font-size: 14pt;
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            }

            QLineEdit, QDateEdit{
                border: 1px solid #ccc;
                border-radius: 8px;
                padding: 5px;
                background: #fff;
            }

            QPushButton {
                background-color: #007aff;
                color: #fff;
                border-radius: 8px;
                padding: 6px 12px;
                border: 1px solid #007aff;
                margin: 5px;
            }

            QPushButton:hover {
                background-color: #005ecb;
            }

            QRadioButton {
                spacing: 5px;
            }

            QScrollArea {
                border: 1px solid #ccc;
                border-radius: 8px;
            }

            QSpinBox:disabled, QSpinBox:!enabled {  # 비활성화된 QSpinBox 스타일
            background: #e6e6e6;
            color: #999;  # 글자색 변경
            }
        """)

    def setupEntityTypeGroup(self, mainLayout):
        entityTypeBox = QGroupBox("엔터티 유형 선택")
        entityTypeLayout = QHBoxLayout()
        entityTypeLayout.setSpacing(5)  # 라디오 버튼 사이의 간격 조정
        self.entityTypeGroup = QButtonGroup(self)
        self.corporateButton = QRadioButton('법인')
        self.individualButton = QRadioButton('개인')
        self.entityTypeGroup.addButton(self.corporateButton)
        self.entityTypeGroup.addButton(self.individualButton)
        entityTypeLayout.addWidget(self.corporateButton)
        entityTypeLayout.addWidget(self.individualButton)
        entityTypeBox.setLayout(entityTypeLayout)
        mainLayout.addWidget(entityTypeBox)


    def setupCalculationTypeGroup(self, mainLayout):
        calculationTypeBox = QGroupBox("계산 유형 선택")
        calculationTypeLayout = QHBoxLayout()
        self.calculationTypeGroup = QButtonGroup(self)
        self.typeAButton = QRadioButton('A유형')
        self.typeBButton = QRadioButton('B유형')
        self.calculationTypeGroup.addButton(self.typeAButton)
        self.calculationTypeGroup.addButton(self.typeBButton)
        self.typeAButton.toggled.connect(self.calculateRemuneration)
        self.typeBButton.toggled.connect(self.calculateRemuneration)
        calculationTypeLayout.addWidget(self.typeAButton)
        calculationTypeLayout.addWidget(self.typeBButton)
        calculationTypeBox.setLayout(calculationTypeLayout)
        self.typeBButton.setChecked(True)
        mainLayout.addWidget(calculationTypeBox)

    def setupCostProgression(self, mainLayout):
        costProgressionBox = QGroupBox("원가계산 누진 선택")
        costProgressionLayout = QHBoxLayout()
        self.costProgressionGroup = QButtonGroup(self)
        self.costProgressionGroup.setExclusive(True)  # 중복 선택 불가
        self.percentages = [0, 10]
        for percent in self.percentages:
            rb = QRadioButton(f"{percent}%")
            self.costProgressionGroup.addButton(rb)
            rb.toggled.connect(self.calculateRemuneration)
            costProgressionLayout.addWidget(rb)
        self.costProgressionGroup.buttons()[0].setChecked(True)  # 첫 번째 라디오 버튼을 디폴트로 선택
        costProgressionBox.setLayout(costProgressionLayout)
        mainLayout.addWidget(costProgressionBox)

    def setupNumPeople(self, mainLayout):
        self.numPeopleLayout = QHBoxLayout()
        self.numPeopleLabel = QLabel("인원수:")
        self.numPeopleSpinBox = QSpinBox()
        self.numPeopleSpinBox.setMinimum(2)
        self.numPeopleSpinBox.setValue(2)

        # QPalette를 사용하여 비활성화된 QSpinBox의 배경색과 텍스트 색상 설정
        palette = self.numPeopleSpinBox.palette()  # QSpinBox의 현재 팔레트 가져오기
        palette.setColor(QPalette.Disabled, QPalette.Base, Qt.gray)  # 배경색 설정
        palette.setColor(QPalette.Disabled, QPalette.Text, Qt.darkGray)  # 텍스트 색상 설정
        self.numPeopleSpinBox.setPalette(palette)

        self.numPeopleSpinBox.setEnabled(False)  # 처음에는 비활성화 상태로 설정
        self.numPeopleSpinBox.valueChanged.connect(self.calculateRemuneration)
        self.numPeopleLayout.addWidget(self.numPeopleLabel)
        self.numPeopleLayout.addWidget(self.numPeopleSpinBox)
        mainLayout.addLayout(self.numPeopleLayout)

    def setupDueDate(self, mainLayout):
        self.dueDateLabel = QLabel("납기일:")
        self.dueDateEdit = QDateEdit()
        self.dueDateEdit.setCalendarPopup(True)  # 캘린더 팝업 활성화
        self.dueDateEdit.setDate(QDate(2024, 3, 31))  # 디폴트 날짜 설정
        self.dueDateEdit.dateChanged.connect(self.calculateRemuneration)  # 날짜 변경 시 계산 메소드 호출
        dueDateLayout = QHBoxLayout()
        dueDateLayout.addWidget(self.dueDateLabel)
        dueDateLayout.addWidget(self.dueDateEdit)
        mainLayout.addLayout(dueDateLayout)

    def setupBaseCalcProgression(self, mainLayout):
        baseCalcProgressionBox = QGroupBox("기본계산 누진 선택")
        baseCalcProgressionLayout = QHBoxLayout()
        self.baseCalcProgressionGroup = QButtonGroup(self)
        self.baseCalcProgressionGroup.setExclusive(False)  # 다중 선택 허용

        self.basePercentages = ['일반', '공동사업자(인원)', '건설업 등(40%)', '외감대상(50%)']
        for i, percent in enumerate(self.basePercentages):
            rb = QRadioButton(percent)
            self.baseCalcProgressionGroup.addButton(rb)
            baseCalcProgressionLayout.addWidget(rb)
            rb.toggled.connect(lambda checked, rb=rb: self.toggleBaseCalcProgression(rb, checked))

        baseCalcProgressionBox.setLayout(baseCalcProgressionLayout)
        mainLayout.addWidget(baseCalcProgressionBox)

    def toggleBaseCalcProgression(self, rb, checked):
        if rb.text() == "일반" and checked:
            for button in self.baseCalcProgressionGroup.buttons():
                if button != rb:
                    button.setChecked(False)
        elif checked:
            self.baseCalcProgressionGroup.buttons()[0].setChecked(False)

        # "일반" 외의 모든 버튼이 비활성화되어 있을 때 "일반" 버튼 자동 체크
        nonGeneralButtonsChecked = any(button.isChecked() for button in self.baseCalcProgressionGroup.buttons() if button.text() != "일반")
        if not nonGeneralButtonsChecked:
            self.baseCalcProgressionGroup.buttons()[0].setChecked(True)


        self.toggleNumPeopleSpinBox()
        self.calculateRemuneration()  # 추가 보수 라벨 업데이트를 위해 호출

    def toggleNumPeopleSpinBox(self):
        # "공동사업자(인원)" 버튼이 체크되어 있으면 인원수 입력란 활성화, 그렇지 않으면 비활성화
        isJointBusinessChecked = self.baseCalcProgressionGroup.buttons()[1].isChecked()
        self.numPeopleSpinBox.setEnabled(isJointBusinessChecked)

    def calculateRemuneration(self):
        if not hasattr(self, 'incomeEdit') or not self.incomeEdit.text():
            return

        if not self.validateInputs():
            return

        # 기본 설정
        entityType = '법인' if self.corporateButton.isChecked() else '개인'
        calculationType = 'B유형' if self.typeBButton.isChecked() else 'A유형'

        income = float(self.incomeEdit.text())
        remuneration = self.calculateBasicRemuneration(income, entityType, calculationType)

        additionalRemuneration = 0
        additionalDetails = ''

        # 추가 보수 계산
        if self.baseCalcProgressionGroup.buttons()[1].isChecked():  # 공동사업자(인원)
            numPeople = self.numPeopleSpinBox.value()
            jointBusinessAddition = remuneration * (numPeople - 1) * 0.2
            additionalRemuneration += jointBusinessAddition
            additionalDetails += f"공동사업자({numPeople}인): +{int(jointBusinessAddition)}\n"

        if self.baseCalcProgressionGroup.buttons()[2].isChecked():  # 건설업 등(40%)
            constructionAddition = remuneration * 0.4
            additionalRemuneration += constructionAddition
            additionalDetails += f"건설업 추가: +{int(constructionAddition)}\n"

        if self.baseCalcProgressionGroup.buttons()[3].isChecked():  # 외감대상(50%)
            auditAddition = remuneration * 0.5
            additionalRemuneration += auditAddition
            additionalDetails += f"외감대상 추가: +{int(auditAddition)}\n"

        # 기본 보수 및 최종 보수 계산 및 라벨 업데이트
        finalRemuneration = remuneration + additionalRemuneration
        details = f'기본보수: {int(remuneration)}\n'
        final_details = f"최종 보수: {int(finalRemuneration)}"

        self.resultLabel.setText(details)

        # 추가 보수 라벨 업데이트
        if additionalRemuneration > 0:
            self.additionalLabel.setText(additionalDetails)
            self.additionalLabel.show()
        else:
            self.additionalLabel.hide()

        self.final_resultLabel.setText(final_details)
        self.finalRemuneration = finalRemuneration
        
        # 원가계산 누진 계산
        selectedCostButton = self.costProgressionGroup.checkedButton()
        if selectedCostButton:
            selectedPercentage = self.percentages[self.costProgressionGroup.buttons().index(selectedCostButton)]
            costProgression = int(finalRemuneration * selectedPercentage / 100)
            self.costProgressionLabel.setText(f'원가계산누진: {costProgression}')

    def calculateBasicRemuneration(self, income, entityType, calculationType):
        if calculationType == 'A유형':
            # A유형 개인
            if entityType == '개인':
                base_amount, rate, min = 300000, 0, 0
                if income >= 1e8:
                    base_amount, rate, min = 300000, 0.0015, 1e8
                if income >= 3e8:
                    base_amount, rate, min = 600000, 0.001, 3e8
                if income >= 5e8:
                    base_amount, rate, min = 800000, 0.0008, 5e8
                if income >= 10e8:
                    base_amount, rate, min = 1200000, 0.0006, 10e8
                if income >= 30e8:
                    base_amount, rate, min = 2400000, 0.0004, 30e8
                if income >= 50e8:
                    base_amount, rate, min = 3200000, 0.0002, 50e8
                if income >= 100e8:
                    base_amount, rate, min = 4200000, 0.00018, 100e8
                if income >= 500e8:
                    base_amount, rate, min = 11400000, 0.00016, 500e8
                if income >= 1000e8:
                    base_amount, rate, min = 19400000, 0.00014, 1000e8

            else:  # A유형 법인
                base_amount, rate, min = 300000, 0, 0
                if income >= 1e8:
                    base_amount, rate, min = 400000, 0.0015, 1e8
                if income >= 3e8:
                    base_amount, rate, min = 700000, 0.001, 3e8
                if income >= 5e8:
                    base_amount, rate, min = 900000, 0.0008, 5e8
                if income >= 10e8:
                    base_amount, rate, min = 1300000, 0.0006, 10e8
                if income >= 30e8:
                    base_amount, rate, min = 2500000, 0.0004, 30e8
                if income >= 50e8:
                    base_amount, rate, min = 3300000, 0.0002, 50e8
                if income >= 100e8:
                    base_amount, rate, min = 4300000, 0.00018, 100e8
                if income >= 500e8:
                    base_amount, rate, min = 11500000, 0.00016, 500e8
                if income >= 1000e8:
                    base_amount, rate, min = 19500000, 0.00014, 1000e8
        else:  # B유형 개인
            if entityType == '개인':
                base_amount, rate, min = 300000, 0, 0
                if income >= 1e8:
                    base_amount, rate, min = 300000, 0.002, 1e8
                if income >= 3e8:
                    base_amount, rate, min = 700000, 0.0016, 3e8
                if income >= 5e8:
                    base_amount, rate, min = 1020000, 0.0014, 5e8
                if income >= 10e8:
                    base_amount, rate, min = 1720000, 0.0012, 10e8
                if income >= 30e8:
                    base_amount, rate, min = 4120000, 0.001, 30e8
                if income >= 50e8:
                    base_amount, rate, min = 6120000, 0.0008, 50e8
                if income >= 100e8:
                    base_amount, rate, min = 10120000, 0.0005, 100e8
                if income >= 500e8:
                    base_amount, rate, min = 30120000, 0.0002, 500e8
                if income >= 1000e8:
                    base_amount, rate, min = 40120000, 0.00008, 1000e8

            else:  # B유형 법인
                base_amount, rate, min = 300000, 0, 0
                if income >= 1e8:
                    base_amount, rate, min = 400000, 0.002, 1e8
                if income >= 3e8:
                    base_amount, rate, min = 800000, 0.0016, 3e8
                if income >= 5e8:
                    base_amount, rate, min = 1120000, 0.0014, 5e8
                if income >= 10e8:
                    base_amount, rate, min = 1820000, 0.0012, 10e8
                if income >= 30e8:
                    base_amount, rate, min = 4220000, 0.001, 30e8
                if income >= 50e8:
                    base_amount, rate, min = 6220000, 0.0008, 50e8
                if income >= 100e8:
                    base_amount, rate, min = 10220000, 0.0005, 100e8
                if income >= 500e8:
                    base_amount, rate, min = 30220000, 0.0002, 500e8
                if income >= 1000e8:
                    base_amount, rate, min = 40220000, 0.00008, 1000e8
        
        if rate >= 0:
            excess_income = income - min
            remuneration = base_amount + excess_income * rate
        else:
            remuneration = base_amount

        remuneration = (remuneration // 10000) * 10000

        return remuneration
    
    def create_excel_file(self):
        if not self.validateInputs():
            return
        company_name = self.companyNameEdit.text()
        if not company_name:
            self.messageLabel.setText("회사명을 입력해주세요.")
            return
        result = int(self.finalRemuneration)  # 소수점 제거
        costProgression = int(self.costProgressionLabel.text().split(': ')[1])  # 소수점 제거

        template_path = self.resource_path('양식.xlsx')
        output_path = f'{company_name}2023년귀속 조정보수청구서.xlsx'

        # 파일이 이미 존재하는지 확인
        if os.path.exists(output_path):
            # 사용자에게 확인 받기
            reply = QMessageBox.question(self, '파일 덮어쓰기 확인', f"'{output_path}' 파일이 이미 존재합니다. 덮어쓰시겠습니까?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.No:
                return  # 사용자가 '아니오'를 선택하면 작업 취소

        # 양식.xlsx 파일 로드
        workbook = load_workbook(filename=template_path)
        sheet = workbook.active

        # 셀에 데이터 쓰기
        sheet['H4'] = datetime.now().date()
        sheet['B5'] = company_name
        sheet['G18'] = result
        sheet['G19'] = costProgression

        dueDate = self.dueDateEdit.date().toString("yyyy-MM-dd")
        sheet['H6'] = f"납기일:{dueDate}"

        # 수정된 내용을 새 파일로 저장 (기존 파일이 있으면 덮어씀)
        workbook.save(filename=output_path)

        self.messageLabel.setText(f"'{output_path}' 파일 생성 완료.")


    def validateInputs(self):
        # 입력란에 이상한거 쓸 때
        try:
            float(self.incomeEdit.text())
        except ValueError:
            self.messageLabel.setText("기준금액에 유효한 숫자를 입력해주세요.")
            return False

        # 메세지라벨 리셋
        self.messageLabel.setText("")
        return True

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Request_Money()
    ex.show()
    sys.exit(app.exec_())