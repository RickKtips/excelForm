// Tell TypeScript that the XLSX variable is available in the global scope
declare var XLSX: any;

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('excel-form') as HTMLFormElement;
    const addTimeButton = document.getElementById('add-time') as HTMLButtonElement;
    const timeFieldsContainer = document.getElementById('time-fields') as HTMLDivElement;
    const fileInput = document.getElementById('excel-file') as HTMLInputElement;
    let timeInputCount = 1;

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

        timeInputContainer.appendChild(newTimeInput);
        timeInputContainer.appendChild(daysOfWeekContainer);
        timeInputContainer.appendChild(errorMessage);
        timeFieldsContainer.appendChild(timeInputContainer);
        timeInputCount++;
    });

    const validateForm = (): boolean => {
        let isValid = true;
        document.querySelectorAll('.error-message').forEach(el => el.textContent = '');
        document.querySelectorAll('.error').forEach(el => el.classList.remove('error'));

        const inputs = form.querySelectorAll('input[required], select[required]') as NodeListOf<HTMLInputElement | HTMLSelectElement>;
        inputs.forEach(input => {
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

            const isTimeMissing = !timeField.value;
            const areDaysMissing = checkboxes.length === 0;

            timeField.classList.toggle('error', isTimeMissing);

            if (isTimeMissing || areDaysMissing) {
                isValid = false;
                let message = '';
                if (isTimeMissing && areDaysMissing) {
                    message = 'Time is required and at least one day must be selected.';
                } else if (isTimeMissing) {
                    message = 'Time is required.';
                } else {
                    message = 'At least one day must be selected.';
                }
                errorMessageElement.textContent = message;
            }
        });

        return isValid;
    };

    form.addEventListener('submit', (event) => {
        event.preventDefault();

        if (!validateForm()) {
            return;
        }

        const file = fileInput.files ? fileInput.files[0] : null;
        if (!file) {
            // This case is handled by the validator, but we keep it as a safeguard.
            alert('Please upload an Excel file.');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target?.result;
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

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

            const rowData = [
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
