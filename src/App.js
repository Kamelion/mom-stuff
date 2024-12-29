import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import Select from 'react-select';
import _ from 'lodash';

function App() {
  const [fileData, setFileData] = useState([]);
  const [operation, setOperation] = useState(null);
  const [errors, setErrors] = useState([]);

  const addAugmentedData = false;
  const addValidation = true;

  const columns = {
    isSeasonal: 'ft seas inv',
    addressees: 'addressees',
    primaryAddressee: 'primary addressee',
    primaryDateToMail: 'primary date to mail',
    primaryBirthday: 'his bday',
    secondaryddressee: 'secondary addressee',
    secondaryDateToMail: 'secondary date to mail',
    secondaryBirthday: 'her bday',
    primaryAddress1: 'best mailing address',
    primaryCity: 'city 1',
    primaryState: 'st 1',
    primaryZip: 'zip 1',
    seasonalStartDate: 'seasonal start date',
    seasonalEndDate: 'seasonal end date',
    seasonalAddress1: 'other address if seasonal',
    seasonalCity: 'city 2',
    seasonalState: 'st 2',
    seasonalZip: 'zip 2'
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet);

      const cleansedData = cleanseData(json);

      if (cleansedData?.length)
        setFileData(cleansedData);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleOperationChange = (selectedOption) => {
    setOperation(selectedOption.value);
  };

  const handleOperation = () => {
    let modifiedData = [...fileData];

    if (operation === 'birthday') {
      modifiedData = getBirthdayDataSet(modifiedData);
    } else if (operation === 'all') {
      modifiedData = getAllDataSet(modifiedData);
    }
    
    generateNewSpreadsheet(modifiedData);
  };

  const cleanseData = (data) => {
    const dataWithWeirdCharactersRemoved = stripWeirdCharacters(data);
    const transformedData = transformColumnNames(dataWithWeirdCharactersRemoved);
    const updatedData = transformSeasonalToBoolean(transformedData);
    const filteredData = removeEmptyAddresseeRows(updatedData);
    const augmentedData = augmentData(filteredData);

    validateData(augmentedData);

    if (errors?.length) return;

    return organizeData(augmentedData);
  };

  const stripWeirdCharacters = (data) => {
    return data.map(obj => {
      const cleanedObj = {};
      for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
          const cleanedKey = key
            .replace(/[^a-zA-Z0-9 ]/g, "")  // Remove non-alphanumeric characters
            .replace(/\s+/g, " ")           // Replace multiple spaces with a single space
            .trim();                        // Trim leading and trailing spaces

          let cleanedValue = obj[key];
          if (cleanedValue === '.....' || obj[key] === '....') {
            cleanedValue = null;
          } else if (typeof cleanedValue === 'string') {
            cleanedValue = cleanedValue.trim();
          }

          cleanedObj[cleanedKey] = cleanedValue;
        }
      }
      return cleanedObj;
    });
  };

  const transformColumnNames = (data) => {
    const reversedMap = _.invert(columns);

    return data.map(row => {
      return _.mapKeys(row, (value, key) => reversedMap[key.toLowerCase()] || key.toLowerCase());
    });
  };

  const transformSeasonalToBoolean = (data) => {
    return data.map(row => {
      row.isSeasonal = row.isSeasonal === 'FT' ? false : true;
      return row;
    });
  };

  const removeEmptyAddresseeRows = (data) => {
    return data.filter(row => row.addressees);
  };

  const validateData = (data) => {
    if (!addValidation) return;

    const validationErrors = [];

    data.map(row => {
      if (!row.addressees)
        validationErrors.push(`Missing addressees data somewhere.`);

      if (row.isSeasonal === '')
        validationErrors.push(`Missing seasonal data for ${row.addressees}.`);

      if (!row.primaryAddressee)
        validationErrors.push(`Missing primary addreessee for ${row.addressees}.`);

      if (!row.primaryBirthday)
        validationErrors.push(`Missing primary birthday for ${row.addressees}.`);

      if (row.secondaryAddressee && !row.secondaryBirthday)
        validationErrors.push(`Missing secondary birthday for ${row.addressees}.`);

      if (!row.primaryAddress1)
        validationErrors.push(`Missing primary address1 for ${row.addressees}.`);

      if (!row.primaryCity)
        validationErrors.push(`Missing primary city for ${row.addressees}.`);

      if (!row.primaryState)
        validationErrors.push(`Missing primary state for ${row.addressees}.`);

      if (!row.primaryZip)
        validationErrors.push(`Missing primary zip for ${row.addressees}.`);

      if (!row.seasonalStartDate || !row.seasonalEndDate)
        validationErrors.push(`Missing valid seasonal date range for ${row.addressees}.`);

      if (row.IsSeasonal) {
        if (!row.seasonalAddress1)
          validationErrors.push(`Seasonal: Missing seasonal address1 for ${row.addressees}.`);

        if (!row.seasonalCity)
          validationErrors.push(`Seasonal: Missing seasonal city for ${row.addressees}.`);
  
        if (!row.seasonalState)
          validationErrors.push(`Seasonal: Missing seasonal state for ${row.addressees}.`);
  
        if (!row.seasonalZip)
          validationErrors.push(`Seasonal: Missing seasonal zip for ${row.addressees}.`);
      }
    });

    setErrors(validationErrors);
  };

  const augmentData = (data) => {
    if (!addAugmentedData) return data;
    
    return data.map(row => {
      const seasonalStartDate = '01-01';
      const seasonalEndDate = '03-31';

      return {...row, seasonalStartDate, seasonalEndDate };
    });
  };

  const organizeData = (data) => {
    return data.map(row => {
      return {
        ...row,
        primaryAddress: `${row.primaryAddress1}, ${row.primaryCity}, ${row.primaryState} ${row.primaryZip}`,
        primaryLabelAddressFormatting: `${row.primaryAddress1}\n${row.primaryCity}, ${row.primaryState}, ${row.primaryZip}`,
        seasonalAddress: row.isSeasonal ? `${row.seasonalAddress1}, ${row.seasonalCity}, ${row.seasonalState} ${row.seasonalZip}` : null,
        seasonalLabelAddressFormatting: row.isSeasonal ? `${row.seasonalAddress1}\n${row.seasonalCity}, ${row.seasonalState}, ${row.seasonalZip}` : null,
      };
    });
  };

  const getBirthdayDataSet = (data) => {
    const separatedBirthdays = splitAddresseesToSeparateRows(data);
    const filteredBirthdays = filterByBirthday(separatedBirthdays);
    return formatBirthdayResult(filteredBirthdays);
  };

  const splitAddresseesToSeparateRows = (data) => {
    const result = [];

    data.map(row => {
      const selectedAddress = row.isSeasonal ? getSelectedAddress(row) : row.primaryLabelAddressFormatting;

      result.push({
        label: `${row.primaryAddressee}\n${selectedAddress}`,
        birthdayMonth: getBirthdayMonth(row.primaryBirthday),
        birthday: formatDateToMMDD(row.primaryBirthday),
        mailingBirthdayCutoffDate: formatDateToMMDD(getMailingCutoffDateFromBirthday(row.primaryBirthday)),
        isSeasonal: row.isSeasonal,
        seasonalStartDate: row.seasonalStartDate,
        seasonalEndDate: row.seasonalEndDate,
        primaryAddress: row.primaryAddress,
        seasonalAddress: row.seasonalAddress,
      });

      if (row.seasonalAddressee)
        result.push({
          label: `${row.seasonalAddressee}\n${selectedAddress}`,
          birthdayMonth: getBirthdayMonth(row.seasonalBirthday),
          birthday: formatDateToMMDD(row.secondaryBirthday),
          mailingBirthdayCutoffDate: formatDateToMMDD(getMailingCutoffDateFromBirthday(row.secondaryBirthday)),
          isSeasonal: row.isSeasonal,
          seasonalStartDate: row.seasonalStartDate,
          seasonalEndDate: row.seasonalEndDate,
          primaryAddress: row.primaryAddress,
          seasonalAddress: row.seasonalAddress,
        });
    });

    return result;
  };

  const getSelectedAddress = ({ primaryLabelAddressFormatting, seasonalLabelAddressFormatting, seasonalStartDate, seasonalEndDate }) => {
    return isDateInRange({ seasonalStartDate, seasonalEndDate }) ? seasonalLabelAddressFormatting : primaryLabelAddressFormatting;
  };

  const isDateInRange = ({ seasonalStartDate, seasonalEndDate }) => {
    const MAILING_BUFFER_IN_DAYS = 7;

    const today = new Date();
    const mailingBufferDate = new Date(today);
    mailingBufferDate.setDate(today.getDate() + MAILING_BUFFER_IN_DAYS);

    const mailingBufferYear = mailingBufferDate.getFullYear();

    const formattedStartDate = new Date(`${mailingBufferYear}-${seasonalStartDate}`);
    const formattedEndDate = new Date(`${mailingBufferYear}-${seasonalEndDate}`);

    if (formattedEndDate < formattedStartDate)
      formattedEndDate.setFullYear(mailingBufferYear + 1);

    return mailingBufferDate >= formattedStartDate && mailingBufferDate <= formattedEndDate;
  };

  const getBirthdayFromExcelValue = (date) => {
    if (!date) return;

    return excelDateToJSDate(date);
  };

  const excelDateToJSDate = (date) => {
    if (!date) return;

    const excelEpoch = new Date(1899, 11, 30); // Excel's epoch starts from Dec 30, 1899
    return new Date(excelEpoch.getTime() + date * 86400000); // Add serial days in milliseconds
  };

  const getBirthdayMonth = (date) => {
    if (!date) return;

    const birthday = getBirthdayFromExcelValue(date);
    return birthday.getMonth() + 1;
  };

  const formatDateToMMDD = (date) => {
    if (!date) return;

    const dateToUse = date instanceof Date && !isNaN(date) ? date : excelDateToJSDate(date);

    const month = ('0' + (dateToUse.getMonth() + 1)).slice(-2);
    const day = ('0' + dateToUse.getDate()).slice(-2);
  
    return `${month}-${day}`;
  }

  const getMailingCutoffDateFromBirthday = (date) => {
    if (!date) return;

    const MAILING_CUTOFF_IN_DAYS = 7;

    const birthday = getBirthdayFromExcelValue(date);
    birthday.setDate(birthday.getDate() - MAILING_CUTOFF_IN_DAYS);
    return birthday;
  };

  const filterByBirthday = (data) => {
    return data.filter(row => row.birthdayMonth === getNextMonth());
  };

  const getNextMonth = () => {
    const today = new Date();
    const currentMonth = today.getMonth() + 1;
    return (currentMonth % 12) + 1;
  };

  const formatBirthdayResult = (data) => {
    const formattedData = formatDataForBetterReadability(data);
    const updatedColumnData = formatResultColumns(formattedData);

    return updatedColumnData;
  };

  const formatDataForBetterReadability = (data) => {
     return data.map(row => {
      row.birthdayMonth = getMonthName(row.birthdayMonth);
      row.isSeasonal = row.isSeasonal ? 'Yes' : 'No';
      return row;
    });
  };

  const getMonthName = (num, locale = 'en-US') => {
    if (!num) return;

    const date = new Date(2000, num - 1);
    return new Intl.DateTimeFormat(locale, { month: 'long' }).format(date);
  };

  const formatResultColumns = (data) => {
    return data.map(row => {
      return _.mapKeys(row, (value, key) => {
        return _.toUpper(_.startCase(key));
      });
    });
  };

  const getAllDataSet = (data) => {
    const result = [];

    data.map(row => {
      const selectedAddress = row.isSeasonal ? getSelectedAddress(row) : row.primaryLabelAddressFormatting;

      result.push({
        label: `${row.addressees}\n${selectedAddress}`,
        isSeasonal: row.isSeasonal ? 'Yes' : 'No',
        seasonalStartDate: row.seasonalStartDate,
        seasonalEndDate: row.seasonalEndDate,
        primaryAddress: row.primaryAddress,
        seasonalAddress: row.seasonalAddress,
      });
    });

    return formatResultColumns(result);
  };

  const generateNewSpreadsheet = (data) => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `Labels`);
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const filename = getFilename();
    downloadFile({ url, filename });
  };

  const getFilename = () => {
    let addonText;

    if (operation === 'birthday') {
      addonText = `${getNextMonthName()} Birthdays`;
    } else {
      addonText = `All`;
    }

    return `Labels for ${addonText} - ${new Date().toJSON()}.xlsx`;
  };

  const getNextMonthName = () => {
    const date = new Date();
    date.setMonth(date.getMonth() + 1);
    return date.toLocaleDateString('default', { month: 'long' });
  }

  const downloadFile = ({ url, filename }) => {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
  };

  const operations = [
    { label: 'Birthday Labels', value: 'birthday' },
    { label: 'Labels for All', value: 'all' },
  ];

  return (
    <div style={{ padding: '20px' }}>
      <h1>Carol's Labels</h1>
      <input type='file' accept='.xlsx, .xls' onChange={handleFileUpload} />
      <Select options={operations} onChange={handleOperationChange} />
      <button onClick={handleOperation} disabled={!operation}>Download</button>

      {errors.length ? (
        <div>
          <h2>Validation Errors</h2>
          <ul style={{ color: 'red' }}>
            {errors.map((error, index) => (
              <li key={index}>{error}</li>
            ))}
          </ul>
        </div>
      ) : (
        <p style={{ color: 'green', display: fileData?.length ? 'block' : 'none' }}>Data is good.</p>
      )}
    </div>
  );
}

export default App;