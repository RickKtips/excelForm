// Tell TypeScript that the XLSX variable is available in the global scope
declare var XLSX: any;

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('excel-form') as HTMLFormElement;
    const addTimeButton = document.getElementById('add-time') as HTMLButtonElement;
    const timeFieldsContainer = document.getElementById('time-fields') as HTMLDivElement;
    const fileInput = document.getElementById('excel-file') as HTMLInputElement;
    const newButton = document.getElementById('new-button') as HTMLButtonElement;
    const tableBody = document.querySelector('#data-table tbody') as HTMLTableSectionElement;
    let timeInputCount = 1;
    let tableData: { [key: string]: any }[] = [];

    document.querySelectorAll('.remove-time').forEach(button => {
        button.addEventListener('click', (event) => {
            const target = event.target as HTMLElement;
            target.closest('.time-input')?.remove();
        });
    });

    addTimeButton.addEventListener('click', () => {
        const timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');

        const newTimeInput = document.createElement('input');
        newTimeInput.type = 'time';
        newTimeInput.name = 'times[]';

        const daysOfWeekContainer = document.createElement('div');
        daysOfWeekContainer.classList.add('days-of-week');
        const days = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'TODOS'];
        days.forEach(day => {
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = day;
            checkbox.name = `days-${timeInputCount}[]`;
            daysOfWeekContainer.appendChild(checkbox);
            daysOfWeekContainer.append(day);
        });

        const errorMessage = document.createElement('span');
        errorMessage.classList.add('error-message');

        const removeButton = document.createElement('button');
        removeButton.type = 'button';
        removeButton.textContent = 'Remove';
        removeButton.addEventListener('click', () => {
            timeInputContainer.remove();
        });

        timeInputContainer.appendChild(newTimeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);
        timeInputContainer.appendChild(removeButton);
        timeFieldsContainer.appendChild(timeInputContainer);
        timeInputCount++;
    });

    const validateNewButton = (): boolean => {
        let isValid = true;
        document.querySelectorAll('.error-message').forEach(el => el.textContent = '');
        document.querySelectorAll('.error').forEach(el => el.classList.remove('error'));

        const inputs = form.querySelectorAll('input[required], select[required]') as NodeListOf<HTMLInputElement | HTMLSelectElement>;
        inputs.forEach(input => {
            if (input.id === 'excel-file') return;
            const errorMessageElement = input.nextElementSibling as HTMLElement;
            if (!input.value.trim()) {
                isValid = false;
                input.classList.add('error');
                if (errorMessageElement) errorMessageElement.textContent = 'This field is required.';
            }
        });

        const timeInputs = document.querySelectorAll('.time-input') as NodeListOf<HTMLDivElement>;
        timeInputs.forEach((timeInput) => {
            const timeField = timeInput.querySelector('input[type="time"]') as HTMLInputElement;
            const checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked') as NodeListOf<HTMLInputElement>;
            const errorMessageElement = timeInput.querySelector('.error-message') as HTMLElement;

            const timeEntered = timeField.value !== '';
            const daysSelected = checkboxes.length > 0;

            timeField.classList.toggle('error', timeEntered && !daysSelected);

            if (timeEntered !== daysSelected) {
                isValid = false;
                let message = '';
                if (timeEntered && !daysSelected) {
                    message = 'At least one day must be selected.';
                } else if (!timeEntered && daysSelected) {
                    message = 'Time is required.';
                }
                errorMessageElement.textContent = message;
            }
        });

        return isValid;
    };

    const renderTable = () => {
        tableBody.innerHTML = '';
        tableData.forEach(rowData => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${rowData['nome']}</td>
                <td>${rowData['url']}</td>
                <td>${rowData['mensagem']}</td>
                <td>${rowData['zoom']}</td>
                <td>${rowData['url_teams']}</td>
                <td>${rowData['resolucao_tela']}</td>
                <td>${rowData['frequencia_tipo']}</td>
                <td>${rowData['times']}</td>
            `;
            tableBody.appendChild(row);
        });
    };

    newButton.addEventListener('click', () => {
        if (!validateNewButton()) {
            return;
        }

        const formData = new FormData(form);
        const newRow: { [key: string]: any } = {};

        formData.forEach((value, key) => {
            if (key !== 'excel-file' && !key.startsWith('times') && !key.startsWith('days')) {
                newRow[key] = value.toString();
            }
        });

        const timeInputs = document.querySelectorAll('.time-input') as NodeListOf<HTMLDivElement>;
        const formattedTimes: string[] = [];

        timeInputs.forEach((timeInput) => {
            const timeField = timeInput.querySelector('input[type="time"]') as HTMLInputElement;
            const checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked') as NodeListOf<HTMLInputElement>;

            if (timeField.value && checkboxes.length > 0) {
                const selectedDays: string[] = [];
                let hasTodos = false;
                checkboxes.forEach(checkbox => {
                    if (checkbox.value === 'TODOS') hasTodos = true;
                    selectedDays.push(checkbox.value);
                });

                if (hasTodos) {
                    formattedTimes.push(`TODOS-${timeField.value}`);
                } else {
                    formattedTimes.push(`${selectedDays.join('-')}-${timeField.value}`);
                }
            }
        });

        newRow['times'] = formattedTimes.join(', ');
        tableData.push(newRow);
        renderTable();

        // Manually reset form fields except for the file input
        const inputs = form.querySelectorAll('input, select');
        inputs.forEach(input => {
            if (input.id !== 'excel-file') {
                const inputElement = input as HTMLInputElement;
                if (inputElement.type === 'checkbox' || inputElement.type === 'radio') {
                    inputElement.checked = false;
                } else {
                    inputElement.value = '';
                }
            }
        });

        // Clear additional time inputs
        const additionalTimeInputs = document.querySelectorAll('.time-input:not(:first-child)');
        additionalTimeInputs.forEach(input => input.remove());
    });

    const validateSubmitButton = (): boolean => {
        let isValid = true;
        document.querySelectorAll('.error-message').forEach(el => el.textContent = '');
        document.querySelectorAll('.error').forEach(el => el.classList.remove('error'));

        if (!fileInput.files || fileInput.files.length === 0) {
            isValid = false;
            fileInput.classList.add('error');
            const errorMessageElement = fileInput.nextElementSibling as HTMLElement;
            if (errorMessageElement) errorMessageElement.textContent = 'This field is required.';
        }

        if (tableData.length === 0) {
            isValid = false;
            alert('Please add at least one row of data.');
        }

        return isValid;
    };

    form.addEventListener('submit', (event) => {
        event.preventDefault();

        if (!validateSubmitButton()) {
            return;
        }

        const file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
            // This is already handled in validateSubmitButton, but we need the file object.
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const dataToAdd = tableData.map(row => [
                row['nome'],
                row['url'],
                row['mensagem'],
                row['zoom'],
                row['url_teams'],
                row['resolucao_tela'],
                row['frequencia_tipo'],
                row['times']
            ]);

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
