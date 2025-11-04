"use strict";
document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('excel-form');
    var addTimeButton = document.getElementById('add-time');
    var timeFieldsContainer = document.getElementById('time-fields');
    var fileInput = document.getElementById('excel-file');
    var timeInputCount = 1;
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
        timeInputContainer.appendChild(newTimeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);
        timeFieldsContainer.appendChild(timeInputContainer);
        timeInputCount++;
    });
    var validateForm = function () {
        var isValid = true;
        document.querySelectorAll('.error-message').forEach(function (el) { return el.textContent = ''; });
        document.querySelectorAll('.error').forEach(function (el) { return el.classList.remove('error'); });
        var inputs = form.querySelectorAll('input[required], select[required]');
        inputs.forEach(function (input) {
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
            var isTimeMissing = !timeField.value;
            var areDaysMissing = checkboxes.length === 0;
            timeField.classList.toggle('error', isTimeMissing);
            if (isTimeMissing || areDaysMissing) {
                isValid = false;
                var message = '';
                if (isTimeMissing && areDaysMissing) {
                    message = 'Time is required and at least one day must be selected.';
                }
                else if (isTimeMissing) {
                    message = 'Time is required.';
                }
                else {
                    message = 'At least one day must be selected.';
                }
                errorMessageElement.textContent = message;
            }
        });
        return isValid;
    };
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        if (!validateForm()) {
            return;
        }
        var file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
            // This case is handled by the validator, but we keep it as a safeguard.
            alert('Please upload an Excel file.');
            return;
        }
        var reader = new FileReader();
        reader.onload = function (e) {
            var _a;
            var data = (_a = e.target) === null || _a === void 0 ? void 0 : _a.result;
            var workbook = XLSX.read(data, { type: 'array' });
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];
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
            var rowData = [
                newRow['field1'],
                newRow['field2'],
                newRow['field3'],
                newRow['field4'],
                newRow['field5'],
                newRow['select-field'],
                newRow['number-field'],
                newRow['times']
            ];
            XLSX.utils.sheet_add_aoa(worksheet, [rowData], { origin: -1 });
            XLSX.writeFile(workbook, file.name);
            alert('The Excel file has been updated and downloaded!');
            form.reset();
        };
        reader.readAsArrayBuffer(file);
    });
});
