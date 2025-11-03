"use strict";
document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('excel-form');
    var addTimeButton = document.getElementById('add-time');
    var timeFieldsContainer = document.getElementById('time-fields');
    var fileInput = document.getElementById('excel-file');
    addTimeButton.addEventListener('click', function () {
        var timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');
        var newTimeInput = document.createElement('input');
        newTimeInput.type = 'time';
        newTimeInput.name = 'times[]';
        timeInputContainer.appendChild(newTimeInput);
        timeFieldsContainer.appendChild(timeInputContainer);
    });
    form.addEventListener('submit', function (event) {
        event.preventDefault();
        var file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
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
            var times = [];
            formData.forEach(function (value, key) {
                if (key === 'times[]') {
                    if (value) {
                        times.push(value.toString());
                    }
                }
                else if (key !== 'excel-file') { // Exclude the file input from the data
                    newRow[key] = value.toString();
                }
            });
            newRow['times'] = times.join(', ');
            // Order the data to match a consistent column order
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
            // Generate and trigger download of the modified file
            XLSX.writeFile(workbook, file.name);
            alert('The Excel file has been updated and downloaded!');
            form.reset();
        };
        reader.readAsArrayBuffer(file);
    });
});
