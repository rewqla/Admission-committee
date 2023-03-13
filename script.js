let studentDataForm = document.getElementById("studentData");
let StudentEmailForm = document.getElementById("studentEmail");

studentDataForm.addEventListener("submit", async function (e) {
    e.preventDefault();
    let studentDataElements = studentDataForm.elements;
    let studentEmailElements = StudentEmailForm.elements;
    if (CheckFieldsIfEmpty(studentDataElements) && CheckPhone(studentDataElements.phone.value)) {
        studentEmailElements.email.value = await Translator(studentDataElements.lastName.value);
        FillEmailAddon(studentDataElements.institute.value);
        studentEmailElements.password.value = GeneratePassword();
    }
});

const isRequired = value => value === '' ? false : true;

function CheckFieldsIfEmpty(elements) {
    let isValid = true;

    Array.from(elements).forEach(element => {
        if (element.tagName.toLowerCase() === 'input') {
            let errorSpan = document.querySelector("span[for=" + element.id + "]");
            if (!isRequired(element.value)) {
                errorSpan.textContent = "Поле має бути не пустим";
                isValid = false;
            }
            else {
                errorSpan.textContent = "";
            }
        }
    });

    return isValid;
}

function CheckPhone(phone) {
    let regexPhone = new RegExp("^([+]?(38))?(0[0-9]{9})$");
    let errorSpan = document.querySelector("span[for=phone]");

    if (regexPhone.test(phone)) {
        errorSpan.textContent = "";
        return true;
    }

    errorSpan.textContent = "Не правильний формат телефону";
    return false;
}

function CheckPasswordLength(password) {
    let errorSpan = document.querySelector("span[for=password]");

    if (password.length >= 8) {
        errorSpan.textContent = "";
        return true;
    }

    errorSpan.textContent = "Пароль має бути не менше 8 символів";
    return false;
}

function FillEmailAddon(insitute) {
    document.getElementById("emailAddon").style = "display: flex!important;";
    document.getElementById("emailAddon").innerText = `_${insitute}${new Date().getFullYear()}@nuwm.edu.ua`
}

function GeneratePassword() {
    const chars = "0123456789abcdefghijklmnopqrstuvwxyz!@#$%^&*()ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var length = 10, password = "";
    for (var i = 0, n = chars.length; i < length; i++) {
        password += chars.charAt(Math.floor(Math.random() * n));
    }
    return password;
}

async function Translator(lastName) {
    let apiUrl = `https://api.mymemory.translated.net/get?q=${lastName.toLowerCase()}&langpair=uk|en`;
    const response = await fetch(apiUrl);
    const result = await response.json();

    console.log(result);

    return result.responseData.translatedText;
}

let workbook, worksheet, fileData = [], filename = "Students.xlsx", sheets = "студенти";

document.getElementById("ChoseExcel").addEventListener("click", function (e) {
    document.getElementById('excelFile').click();
});

document.getElementById("CreateExcel").addEventListener("click", function (e) {
    workbook = XLSX.utils.book_new();
    workbook.SheetNames.push(sheets);
    HideExcelSection("", "створено");
    fileData.push(["Повне імя студента", "Телефон", "Пошта", "Пароль"]);
});

document.getElementById("excelFile").addEventListener("change", async function (e) {
    const file = e.target.files[0];
    if (file.name.includes("xls") || file.name.includes("xlsx")) {
        HideExcelSection(file.name, "вибраний");

        const data = await file.arrayBuffer();
        workbook = XLSX.read(data);
        sheets = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheets];
        fileData = LoadExcelData(worksheet);
        filename = file.name;
    }
    else {
        document.getElementById("ExcelText").innerText = "Виберіть excel файл!";
    }
});

document.getElementById("AddStudent").addEventListener("click", function (e) {
    if (workbook) {
        let studentDataElements = studentDataForm.elements;
        let studentEmailElements = StudentEmailForm.elements;
        if (CheckFieldsIfEmpty(studentDataElements) && CheckFieldsIfEmpty(studentEmailElements)
            && CheckPhone(studentDataElements.phone.value) && CheckPasswordLength(studentEmailElements.password.value)) {
            fileData.push(GetStudentData(studentDataElements, studentEmailElements));
        }
    }
    else {
        document.getElementById("ExcelText").innerText = "Виберіть або створіть файл!";
    }
});

document.getElementById("DownloadExcel").addEventListener("click", function (e) {
    worksheet = XLSX.utils.aoa_to_sheet(fileData);
    workbook.Sheets[sheets] = worksheet;

    XLSX.writeFile(workbook, filename);
});

function LoadExcelData(worksheet) {
    let range = XLSX.utils.decode_range(worksheet["!ref"])

    var data = [];
    for (let row = range.s.r; row <= range.e.r; row++) {
        let i = data.length;
        data.push([]);
        for (let col = range.s.c; col <= range.e.c; col++) {
            let cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
            data[i].push(cell.v);
        }
    }
    console.log(data);
    return data;
}

function HideExcelSection(fileName, action) {
    document.getElementById("ExcelText").innerText = `Файл ${fileName} ${action}`;
    document.getElementById("choseExcelType").classList.add("d-none")
    document.getElementById("DownloadExcel").classList.remove("d-none");
}

function GetStudentData(firstForm, secondForm) {
    let fullName = `${firstForm.firstName.value} ${firstForm.middleName.value} ${firstForm.lastName.value}`;
    document.getElementById("ActionInfo").innerText = `Студента ${fullName} додано`;
    return [fullName, firstForm.phone.value, secondForm.email.value, secondForm.password.value];
}