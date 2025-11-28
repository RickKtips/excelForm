// Tell TypeScript that the XLSX variable is available in the global scope
declare var XLSX: any;

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('excel-form') as HTMLFormElement;
    const addTimeButton = document.getElementById('add-time') as HTMLButtonElement;
    const timeFieldsContainer = document.getElementById('time-fields') as HTMLDivElement;
    const frequenciaTipo = document.getElementById('frequencia_tipo') as HTMLSelectElement;
    const fileInput = document.getElementById('excel-file') as HTMLInputElement;
    const newButton = document.getElementById('new-button') as HTMLButtonElement;
    const tableBody = document.querySelector('#data-table tbody') as HTMLTableSectionElement;
    let timeInputCount = 1;
    let tableData: { [key: string]: any }[] = [];

    const createTimeInputRow = (index: number, type: string) => {
        const timeInputContainer = document.createElement('div');
        timeInputContainer.classList.add('time-input');

        let timeInput: HTMLInputElement | HTMLSelectElement;

        if (type === 'constant') {
            timeInput = document.createElement('select');
            timeInput.name = `times[]`;
            const options = ['15min', '1h', '2h', '4h', '8h'];
            options.forEach(optionValue => {
                const option = document.createElement('option');
                option.value = optionValue;
                option.textContent = optionValue;
                timeInput.appendChild(option);
            });
        } else {
            timeInput = document.createElement('input');
            timeInput.type = 'time';
            timeInput.name = `times[]`;
        }

        const daysOfWeekContainer = document.createElement('div');
        daysOfWeekContainer.classList.add('days-of-week');
        const days = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'TODOS'];
        days.forEach(day => {
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = day;
            checkbox.name = `days-${index}[]`;
            daysOfWeekContainer.appendChild(checkbox);
            daysOfWeekContainer.append(day);
        });

        const errorMessage = document.createElement('span');
        errorMessage.classList.add('error-message');

        timeInputContainer.appendChild(timeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);

        if (index > 0) {
            const deleteButton = document.createElement('button');
            deleteButton.type = 'button';
            deleteButton.classList.add('delete-time');
            deleteButton.textContent = 'Delete';
            timeInputContainer.appendChild(deleteButton);
        }

        return timeInputContainer;
    };

    addTimeButton.addEventListener('click', () => {
        const selectedType = frequenciaTipo.value;
        const newTimeInput = createTimeInputRow(timeInputCount, selectedType);
        timeFieldsContainer.appendChild(newTimeInput);
        timeInputCount++;
    });

    timeFieldsContainer.addEventListener('click', (event) => {
        const target = event.target as HTMLElement;
        if (target.classList.contains('delete-time')) {
            const timeInput = target.closest('.time-input');
            if (timeInput) {
                timeInput.remove();
            }
        }
    });

    const updateTimeInput = () => {
        const selectedValue = frequenciaTipo.value;
        timeFieldsContainer.innerHTML = ''; // Clear existing fields
        const initialTimeInput = createTimeInputRow(0, selectedValue);
        timeFieldsContainer.appendChild(initialTimeInput);
        timeInputCount = 1; // Reset count
        addTimeButton.style.display = 'block';
    };

    frequenciaTipo.addEventListener('change', updateTimeInput);

    // Initialize with one time input row
    updateTimeInput();


    const generateConstantTimes = (interval: string, days: string[]): string[] => {
        const times: string[] = [];
        const value = parseInt(interval);

        let intervalInMinutes = 0;
        if (interval.includes('h')) {
            intervalInMinutes = value * 60;
        } else if (interval.includes('min')) {
            intervalInMinutes = value;
        }

        if (intervalInMinutes === 0) {
            return [];
        }

        for (let m = 0; m < 24 * 60; m += intervalInMinutes) {
            const hour = Math.floor(m / 60).toString().padStart(2, '0');
            const minute = (m % 60).toString().padStart(2, '0');
            times.push(`${days.join('-')}-${hour}:${minute}`);
        }
        return times;
    };

    const validateForm = (): boolean => {
        let isValid = true;
        document.querySelectorAll('.error-message').forEach(el => el.textContent = '');
        document.querySelectorAll('.error').forEach(el => el.classList.remove('error'));

        const requiredFields = [
            'nome', 'url', 'mensagem', 'zoom', 'url_teams',
            'resolucao_tela', 'frequencia_tipo'
        ];

        requiredFields.forEach(fieldId => {
            const input = document.getElementById(fieldId) as HTMLInputElement | HTMLSelectElement;
            const errorMessageElement = input.nextElementSibling as HTMLElement;
            if (!input.value.trim()) {
                isValid = false;
                input.classList.add('error');
                if (errorMessageElement) errorMessageElement.textContent = 'This field is required.';
            }
        });

        const timeInputs = document.querySelectorAll('.time-input') as NodeListOf<HTMLDivElement>;
        timeInputs.forEach((timeInput) => {
            const timeField = timeInput.querySelector('input[type="time"], select') as HTMLInputElement | HTMLSelectElement;
            const checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked') as NodeListOf<HTMLInputElement>;
            const errorMessageElement = timeInput.querySelector('.error-message') as HTMLElement;

            const timeEntered = timeField.value !== '';
            const daysSelected = checkboxes.length > 0;

            if (timeEntered && !daysSelected) {
                isValid = false;
                timeField.classList.add('error');
                if (errorMessageElement) errorMessageElement.textContent = 'At least one day must be selected.';
            } else if (!timeEntered && daysSelected) {
                isValid = false;
                timeField.classList.add('error');
                if (errorMessageElement) errorMessageElement.textContent = 'Time is required.';
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

    const resetForm = () => {
        const selectedFile = fileInput.files ? fileInput.files[0] : null;
        form.reset();
        if (selectedFile) {
            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(selectedFile);
            fileInput.files = dataTransfer.files;
        }


        const timeInputs = timeFieldsContainer.querySelectorAll('.time-input');
        timeInputs.forEach((timeInput, index) => {
            if (index > 0) {
                timeInput.remove();
            }
        });

        document.querySelectorAll('.error-message').forEach(el => el.textContent = '');
        document.querySelectorAll('.error').forEach(el => el.classList.remove('error'));
    };

    newButton.addEventListener('click', () => {
        if (!validateForm()) {
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
        let formattedTimes: string[] = [];

        timeInputs.forEach((timeInput) => {
            const timeField = timeInput.querySelector('input[type="time"], select') as HTMLInputElement | HTMLSelectElement;
            const checkboxes = timeInput.querySelectorAll('input[type="checkbox"]:checked') as NodeListOf<HTMLInputElement>;

            if (timeField.value && checkboxes.length > 0) {
                const selectedDays: string[] = [];
                let hasTodos = false;
                checkboxes.forEach(checkbox => {
                    if (checkbox.value === 'TODOS') hasTodos = true;
                    selectedDays.push(checkbox.value);
                });

                if (frequenciaTipo.value === 'constant') {
                    if (hasTodos) {
                        formattedTimes = formattedTimes.concat(generateConstantTimes(timeField.value, ['TODOS']));
                    } else {
                        formattedTimes = formattedTimes.concat(generateConstantTimes(timeField.value, selectedDays));
                    }
                } else {
                    if (hasTodos) {
                        formattedTimes.push(`TODOS-${timeField.value}`);
                    } else {
                        formattedTimes.push(`${selectedDays.join('-')}-${timeField.value}`);
                    }
                }
            }
        });

        newRow['times'] = formattedTimes.join(', ');
        tableData.push(newRow);
        renderTable();

        resetForm();
    });

    form.addEventListener('submit', (event) => {
        event.preventDefault();

        const file = fileInput.files ? fileInput.files[0] : null;
        const fileError = document.getElementById('file-error') as HTMLElement;

        if (tableData.length === 0) {
            alert('Please add at least one row of data.');
            return;
        }

        if (!file) {
            fileError.textContent = 'Please upload an Excel file.';
            fileInput.classList.add('error');
            return;
        } else {
            fileError.textContent = '';
            fileInput.classList.remove('error');
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
