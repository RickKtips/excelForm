"use strict";
document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('excel-form');
    var addTimeButton = document.getElementById('add-time');
    var timeFieldsContainer = document.getElementById('time-fields');
    var fileInput = document.getElementById('excel-file');
    var newButton = document.getElementById('new-button');
    var tableBody = document.querySelector('#data-table tbody');
    var timeInputCount = 1;
    var tableData = [];
    document.querySelectorAll('.remove-time').forEach(function (button) {
        button.addEventListener('click', function (event) {
            var _a;
            var target = event.target;
            (_a = target.closest('.time-input')) === null || _a === void 0 ? void 0 : _a.remove();
        });
    });
    addTimeButton.addEventListener('click', function () {
        var timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');
        var newTimeInput = document.createElement('input');
        newTimeInput.type = 'time';
        newTimeInput.name = 'times[]';
        var daysOfWeekContainer = document.createElement('div');
        daysOfWeekContainer.classList.add('days-of-week');
        var days = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'TODOS'];
        days.forEach(function (day) {
            var checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = day;
            checkbox.name = "days-".concat(timeInputCount, "[]");
            daysOfWeekContainer.appendChild(checkbox);
            daysOfWeekContainer.append(day);
        });
        var errorMessage = document.createElement('span');
        errorMessage.classList.add('error-message');
        var removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.textContent = 'Remove';
        removeButton.addEventListener('click', function () {
            timeInputContainer.remove();
        });
        timeInputContainer.appendChild(newTimeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);
        timeInputContainer.appendChild(removeButton);
        timeFieldsContainer.appendChild(timeInputContainer);
        timeInputCount++;
    });
    var validateNewButton = function () {
        var isValid = true;
        document.querySelectorAll('.error-message').forEach(function (el) { return el.textContent = ''; });
        document.querySelectorAll('.error').forEach(function (el) { return el.classList.remove('error'); });
        var inputs = form.querySelectorAll('input[required], select[required]');
        inputs.forEach(function (input) {
            if (input.id === 'excel-file')
                return;
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
            var timeField = timeInput.querySelector('input[type="time"]');
            var checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked');
            var errorMessageElement = timeInput.querySelector('.error-message');
            var timeEntered = timeField.value !== '';
            var daysSelected = checkboxes.length > 0;
            timeField.classList.toggle('error', timeEntered && !daysSelected);
            if (timeEntered !== daysSelected) {
                isValid = false;
                var message = '';
                if (timeEntered && !daysSelected) {
                    message = 'At least one day must be selected.';
                }
                else if (!timeEntered && daysSelected) {
                    message = 'Time is required.';
                }
                errorMessageElement.textContent = message;
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
    newButton.addEventListener('click', function () {
        if (!validateNewButton()) {
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
            var timeField = timeInput.querySelector('input[type="time"]');
            var checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked');
            if (timeField.value && checkboxes.length > 0) {
                var selectedDays_1 = [];
                var hasTodos_1 = false;
                checkboxes.forEach(function (checkbox) {
                    if (checkbox.value === 'TODOS')
                        hasTodos_1 = true;
                    selectedDays_1.push(checkbox.value);
                });
                if (hasTodos_1) {
                    formattedTimes.push("TODOS-".concat(timeField.value));
                }
                else {
                    formattedTimes.push("".concat(selectedDays_1.join('-'), "-").concat(timeField.value));
                }
            }
        });
        newRow['times'] = formattedTimes.join(', ');
        tableData.push(newRow);
        renderTable();
        // Manually reset form fields except for the file input
        var inputs = form.querySelectorAll('input, select');
        inputs.forEach(function (input) {
            if (input.id !== 'excel-file') {
                var inputElement = input;
                if (inputElement.type === 'checkbox' || inputElement.type === 'radio') {
                    inputElement.checked = false;
                }
                else {
                    inputElement.value = '';
                }
            }
        });
        // Clear additional time inputs
        var additionalTimeInputs = document.querySelectorAll('.time-input:not(:first-child)');
        additionalTimeInputs.forEach(function (input) { return input.remove(); });
    });
    var validateSubmitButton = function () {
        var isValid = true;
        document.querySelectorAll('.error-message').forEach(function (el) { return el.textContent = ''; });
        document.querySelectorAll('.error').forEach(function (el) { return el.classList.remove('error'); });
        if (!fileInput.files || fileInput.files.length === 0) {
            isValid = false;
            fileInput.classList.add('error');
            var errorMessageElement = fileInput.nextElementSibling;
            if (errorMessageElement)
                errorMessageElement.textContent = 'This field is required.';
        }
        if (tableData.length === 0) {
            isValid = false;
            alert('Please add at least one row of data.');
        }
        return isValid;
    };
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        if (!validateSubmitButton()) {
            return;
        }
        var file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
            // This is already handled in validateSubmitButton, but we need the file object.
            return;
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
