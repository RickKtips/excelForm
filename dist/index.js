"use strict";
document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('excel-form');
    var addTimeButton = document.getElementById('add-time');
    var timeFieldsContainer = document.getElementById('time-fields');
    var frequenciaTipo = document.getElementById('frequencia_tipo');
    var fileInput = document.getElementById('excel-file');
    var newButton = document.getElementById('new-button');
    var tableBody = document.querySelector('#data-table tbody');
    var timeInputCount = 1;
    var tableData = [];
    var createTimeInputRow = function (index, type) {
        var timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');
        var timeInput;
        if (type === 'constant') {
            timeInput = document.createElement('select');
            timeInput.name = "times[]";
            var options = ['15min', '1h', '2h', '4h', '8h'];
            options.forEach(function (optionValue) {
                var option = document.createElement('option');
                option.value = optionValue;
                option.textContent = optionValue;
                timeInput.appendChild(option);
            });
        }
        else {
            timeInput = document.createElement('input');
            timeInput.type = 'time';
            timeInput.name = "times[]";
        }
        var daysOfWeekContainer = document.createElement('div');
        daysOfWeekContainer.classList.add('days-of-week');
        var days = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'TODOS'];
        days.forEach(function (day) {
            var checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = day;
            checkbox.name = "days-".concat(index, "[]");
            daysOfWeekContainer.appendChild(checkbox);
            daysOfWeekContainer.append(day);
        });
        var errorMessage = document.createElement('span');
        errorMessage.classList.add('error-message');
        timeInputContainer.appendChild(timeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);
        if (index > 0) {
            var deleteButton = document.createElement('button');
            deleteButton.type = 'button';
            deleteButton.classList.add('delete-time');
            deleteButton.textContent = 'Delete';
            timeInputContainer.appendChild(deleteButton);
        }
        return timeInputContainer;
    };
    addTimeButton.addEventListener('click', function () {
        var selectedType = frequenciaTipo.value;
        var newTimeInput = createTimeInputRow(timeInputCount, selectedType);
        timeFieldsContainer.appendChild(newTimeInput);
        timeInputCount++;
    });
    timeFieldsContainer.addEventListener('click', function (event) {
        var target = event.target;
        if (target.classList.contains('delete-time')) {
            var timeInput = target.closest('.time-input');
            if (timeInput) {
                timeInput.remove();
            }
        }
    });
    var updateTimeInput = function () {
        var selectedValue = frequenciaTipo.value;
        timeFieldsContainer.innerHTML = ''; // Clear existing fields
        var initialTimeInput = createTimeInputRow(0, selectedValue);
        timeFieldsContainer.appendChild(initialTimeInput);
        timeInputCount = 1; // Reset count
        addTimeButton.style.display = 'block';
    };
    frequenciaTipo.addEventListener('change', updateTimeInput);
    // Initialize with one time input row
    updateTimeInput();
    var generateConstantTimes = function (interval, days) {
        var times = [];
        var value = parseInt(interval);
        var intervalInMinutes = 0;
        if (interval.includes('h')) {
            intervalInMinutes = value * 60;
        }
        else if (interval.includes('min')) {
            intervalInMinutes = value;
        }
        if (intervalInMinutes === 0) {
            return [];
        }
        for (var m = 0; m < 24 * 60; m += intervalInMinutes) {
            var hour = Math.floor(m / 60).toString().padStart(2, '0');
            var minute = (m % 60).toString().padStart(2, '0');
            times.push("".concat(days.join('-'), "-").concat(hour, ":").concat(minute));
        }
        return times;
    };
    var validateForm = function () {
        var isValid = true;
        document.querySelectorAll('.error-message').forEach(function (el) { return el.textContent = ''; });
        document.querySelectorAll('.error').forEach(function (el) { return el.classList.remove('error'); });
        var requiredFields = [
            'nome', 'url', 'mensagem', 'zoom', 'url_teams',
            'resolucao_tela', 'frequencia_tipo'
        ];
        requiredFields.forEach(function (fieldId) {
            var input = document.getElementById(fieldId);
            var errorMessageElement = input.nextElementSibling;
            if (!input.value.trim()) {
                isValid = false;
                input.classList.add('error');
                if (errorMessageElement)
                    errorMessageElement.textContent = 'This field is required.';
            }
        });
        var timeInputs = document.querySelectorAll('.time-input');
        timeInputs.forEach(function (timeInput) {
            var timeField = timeInput.querySelector('input[type="time"], select');
            var checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked');
            var errorMessageElement = timeInput.querySelector('.error-message');
            var timeEntered = timeField.value !== '';
            var daysSelected = checkboxes.length > 0;
            if (timeEntered && !daysSelected) {
                isValid = false;
                timeField.classList.add('error');
                if (errorMessageElement)
                    errorMessageElement.textContent = 'At least one day must be selected.';
            }
            else if (!timeEntered && daysSelected) {
                isValid = false;
                timeField.classList.add('error');
                if (errorMessageElement)
                    errorMessageElement.textContent = 'Time is required.';
            }
        });
        return isValid;
    };
    var renderTable = function () {
        tableBody.innerHTML = '';
        tableData.forEach(function (rowData) {
            var row = document.createElement('tr');
            row.innerHTML = "\n                <td>".concat(rowData['nome'], "</td>\n                <td>").concat(rowData['url'], "</td>\n                <td>").concat(rowData['mensagem'], "</td>\n                <td>").concat(rowData['zoom'], "</td>\n                <td>").concat(rowData['url_teams'], "</td>\n                <td>").concat(rowData['resolucao_tela'], "</td>\n                <td>").concat(rowData['frequencia_tipo'], "</td>\n                <td>").concat(rowData['times'], "</td>\n            ");
            tableBody.appendChild(row);
        });
    };
    var resetForm = function () {
        var selectedFile = fileInput.files ? fileInput.files[0] : null;
        form.reset();
        if (selectedFile) {
            var dataTransfer = new DataTransfer();
            dataTransfer.items.add(selectedFile);
            fileInput.files = dataTransfer.files;
        }
        var timeInputs = timeFieldsContainer.querySelectorAll('.time-input');
        timeInputs.forEach(function (timeInput, index) {
            if (index > 0) {
                timeInput.remove();
            }
        });
        document.querySelectorAll('.error-message').forEach(function (el) { return el.textContent = ''; });
        document.querySelectorAll('.error').forEach(function (el) { return el.classList.remove('error'); });
    };
    newButton.addEventListener('click', function () {
        if (!validateForm()) {
            return;
        }
        var formData = new FormData(form);
        var newRow = {};
        formData.forEach(function (value, key) {
            if (key !== 'excel-file' && !key.startsWith('times') && !key.startsWith('days')) {
                newRow[key] = value.toString();
            }
        });
        var timeInputs = document.querySelectorAll('.time-input');
        var formattedTimes = [];
        timeInputs.forEach(function (timeInput) {
            var timeField = timeInput.querySelector('input[type="time"], select');
            var checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked');
            if (timeField.value && checkboxes.length > 0) {
                var selectedDays_1 = [];
                var hasTodos_1 = false;
                checkboxes.forEach(function (checkbox) {
                    if (checkbox.value === 'TODOS')
                        hasTodos_1 = true;
                    selectedDays_1.push(checkbox.value);
                });
                if (frequenciaTipo.value === 'constant') {
                    if (hasTodos_1) {
                        formattedTimes = formattedTimes.concat(generateConstantTimes(timeField.value, ['TODOS']));
                    }
                    else {
                        formattedTimes = formattedTimes.concat(generateConstantTimes(timeField.value, selectedDays_1));
                    }
                }
                else {
                    if (hasTodos_1) {
                        formattedTimes.push("TODOS-".concat(timeField.value));
                    }
                    else {
                        formattedTimes.push("".concat(selectedDays_1.join('-'), "-").concat(timeField.value));
                    }
                }
            }
        });
        newRow['times'] = formattedTimes.join(', ');
        tableData.push(newRow);
        renderTable();
        resetForm();
    });
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        var file = fileInput.files ? fileInput.files[0] : null;
        var fileError = document.getElementById('file-error');
        if (tableData.length === 0) {
            alert('Please add at least one row of data.');
            return;
        }
        if (!file) {
            fileError.textContent = 'Please upload an Excel file.';
            fileInput.classList.add('error');
            return;
        }
        else {
            fileError.textContent = '';
            fileInput.classList.remove('error');
        }
        var reader = new FileReader();
        reader.onload = function (e) {
            var _a;
            var data = (_a = e.target) === null || _a === void 0 ? void 0 : _a.result;
            var workbook = XLSX.read(data, { type: 'array' });
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];
            var dataToAdd = tableData.map(function (row) { return [
                row['nome'],
                row['url'],
                row['mensagem'],
                row['zoom'],
                row['url_teams'],
                row['resolucao_tela'],
                row['frequencia_tipo'],
                row['times']
            ]; });
            XLSX.utils.sheet_add_aoa(worksheet, dataToAdd, { origin: -1 });
            XLSX.writeFile(workbook, file.name);
            alert('The Excel file has been updated and downloaded!');
            form.reset();
            tableData = [];
            renderTable();
        };
        reader.readAsArrayBuffer(file);
    });
});
