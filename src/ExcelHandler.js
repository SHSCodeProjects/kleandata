import React, { useState, useRef } from 'react';
import ExcelJS from 'exceljs';
import axios from 'axios';
import {
  Button, IconButton, Tabs, Tab, Box, Paper, Table, TableHead, TableRow, TableCell, TableBody, Tooltip, TextField, Dialog, DialogTitle, DialogContent, DialogActions, Select, MenuItem,
} from '@mui/material';
import DeleteIcon from '@mui/icons-material/Delete';
import MergeTypeIcon from '@mui/icons-material/MergeType';
import AddCircleIcon from '@mui/icons-material/AddCircle';
import SearchIcon from '@mui/icons-material/Search';

// Check if running in Electron
const isElectron = () => {
  return window && window.process && window.process.type;
};

// Import shell from electron if running in Electron
let shell;
if (isElectron()) {
  shell = window.require('electron').shell;
}

const ExcelHandler = () => {
  const [workbook, setWorkbook] = useState(null);
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(0);
  const [editCell, setEditCell] = useState({ rowIndex: null, columnKey: null, value: '' });
  const [newRow, setNewRow] = useState({});
  const [mergeDialogOpen, setMergeDialogOpen] = useState(false);
  const [currentRowIndex, setCurrentRowIndex] = useState(null);
  const [targetSheetIndex, setTargetSheetIndex] = useState(null);
  const [manualSelectDialogOpen, setManualSelectDialogOpen] = useState(false);
  const [manualSelectOptions, setManualSelectOptions] = useState([]);
  const [selectedManualRow, setSelectedManualRow] = useState(null);
  const [undoStack, setUndoStack] = useState([]);
  const [isUndoDisabled, setIsUndoDisabled] = useState(true);
  const [searchText, setSearchText] = useState('');
  const tableRef = useRef(null);

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = async (e) => {
      const buffer = e.target.result;
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      setWorkbook(wb);
      loadSheetData(wb.worksheets[0]);
    };
    reader.readAsArrayBuffer(file);
  };

const loadSheetData = (sheet) => {
  const headers = sheet.getRow(1).values.slice(1); // Ignore first element
  const jsonData = sheet.getSheetValues().slice(2).map((row) => {
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header] = row[i + 1];
    });
    return rowData;
  }).filter(row => Object.values(row).some(cell => cell !== null && cell !== undefined && cell !== '')); // Filter out empty rows
  setData(jsonData);
  setFilteredData(jsonData);
};


  const handleSheetChange = (event, newValue) => {
    setSelectedSheet(newValue);
    const sheet = workbook.worksheets[newValue];
    loadSheetData(sheet);
  };

  const deleteRow = async (index) => {
    console.log('Deleting row:', index);
    const rowToDelete = filteredData[index];
    const fullName = `${rowToDelete.firstName} ${rowToDelete.lastName}`.trim().toLowerCase();

    if (!fullName) {
      console.error('Full name is undefined or empty. Row to delete:', rowToDelete);
      return;
    }

    console.log('Full name of the row to delete:', fullName);

    const updatedData = data.filter((_, i) => i !== index);
    setUndoStack([...undoStack, { type: 'delete', data: rowToDelete, index, sheetIndex: selectedSheet }]);
    setData(updatedData);
    setFilteredData(updatedData);

    const sheet = workbook.worksheets[selectedSheet];
    sheet.spliceRows(index + 2, 1); // Adjust for 1-based indexing
    console.log('Deleted row from current sheet');

    const deletedInTarget = await synchronizeDeleteAcrossTabs(fullName, selectedSheet);
    if (!deletedInTarget) {
      console.log('Record not found in target sheet. Deleted from parent sheet only.');
    }

    setIsUndoDisabled(false);
  };

  const synchronizeDeleteAcrossTabs = async (fullName, originatingSheetIndex) => {
    console.log('Synchronizing deletion across tabs for:', fullName, 'from sheet index:', originatingSheetIndex);

    const targetSheetIndex = originatingSheetIndex % 2 === 0 ? originatingSheetIndex + 1 : originatingSheetIndex - 1;
    if (targetSheetIndex >= workbook.worksheets.length || targetSheetIndex < 0) {
      console.log('No corresponding sheet to synchronize with.');
      return;
    }

    const targetSheet = workbook.worksheets[targetSheetIndex];
    let rowIndexToDelete = [];

    const targetHeaders = targetSheet.getRow(1).values.slice(1);
    const fullNameIndex = targetHeaders.indexOf('fullName') + 1;

    if (fullNameIndex === 0) {
      console.error('fullName column not found in target sheet');
      return;
    }

    targetSheet.eachRow((row, rowIndex) => {
      const rowFullName = row.getCell(fullNameIndex).value?.toString().trim().toLowerCase();
      if (rowFullName === fullName) {
        console.log(`Found matching row in target sheet at index: ${rowIndex} with full name: ${rowFullName}`);
        rowIndexToDelete.push(rowIndex);
      }
    });

    if (rowIndexToDelete.length === 0) {
      console.log('No matching row found in target sheet');
      return false;
    }

    console.log('Row indices to delete in target sheet:', rowIndexToDelete);

    rowIndexToDelete.reverse().forEach((rowIndex) => {
      const deletedRow = targetSheet.getRow(rowIndex);
      console.log(`Deleting row at index: ${rowIndex} in target sheet with data:`, deletedRow.values);
      targetSheet.spliceRows(rowIndex, 1);
      console.log('Deleted row at index:', rowIndex, 'in target sheet');
    });

    return true;
  };

  const handleDeleteRow = (index) => {
    deleteRow(index);
  };

  const handleMergeClick = (rowIndex) => {
    setCurrentRowIndex(rowIndex);
    setMergeDialogOpen(true);
  };

  const handleMerge = async () => {
    const selectedRow = data[currentRowIndex];
    if (!selectedRow) return;

    const targetSheet = workbook.worksheets[targetSheetIndex];
    if (!targetSheet) return;

    const targetHeaders = targetSheet.getRow(1).values.slice(1);
    const targetFirstNameIndex = targetHeaders.indexOf('firstName') + 1;
    const targetLastNameIndex = targetHeaders.indexOf('lastName') + 1;

    let matchingRow;
    targetSheet.eachRow((row) => {
      const rowFirstName = row.getCell(targetFirstNameIndex).value?.toString().toLowerCase().trim();
      const rowLastName = row.getCell(targetLastNameIndex).value?.toString().toLowerCase().trim();
      const selectedFullNameParts = selectedRow.fullName.toLowerCase().split(' ');
      const selectedFirstName = selectedFullNameParts[0].trim();
      const selectedLastName = selectedFullNameParts[selectedFullNameParts.length - 1].trim();

      if (rowFirstName === selectedFirstName && rowLastName === selectedLastName) {
        matchingRow = row;
      }
    });

    if (!matchingRow) {
      const options = [];
      targetSheet.eachRow((row) => {
        const firstName = row.getCell(targetFirstNameIndex).value;
        const lastName = row.getCell(targetLastNameIndex).value;
        options.push({ row, firstName, lastName });
      });
      setManualSelectOptions(options);
      setManualSelectDialogOpen(true);
      return;
    }

    await mergeRows(matchingRow, selectedRow, targetHeaders, targetSheet);
    setMergeDialogOpen(false);
  };

  const mergeRows = async (matchingRow, selectedRow, targetHeaders, targetSheet) => {
    const firstEmptyCol = matchingRow.cellCount + 1;
    const originalRowValues = { ...matchingRow.values };
    Object.keys(selectedRow).forEach((key) => {
      if (key !== 'firstName' && key !== 'lastName' && key !== 'fullName') {
        const columnIndex = targetHeaders.indexOf(key) + 1;
        if (columnIndex > 0) {
          matchingRow.getCell(columnIndex).value = selectedRow[key];
        } else {
          const newColumnIndex = firstEmptyCol + targetHeaders.length;
          targetSheet.getRow(1).getCell(newColumnIndex).value = key;
          matchingRow.getCell(newColumnIndex).value = selectedRow[key];
          targetHeaders.push(key);
        }
      }
    });

    setUndoStack([...undoStack, { type: 'merge', originalRowValues, newRowValues: matchingRow.values, sheetIndex: targetSheetIndex }]);

    const updatedData = data.filter((_, i) => i !== currentRowIndex);
    setData(updatedData);
    setFilteredData(updatedData);
    workbook.worksheets[selectedSheet].spliceRows(currentRowIndex + 2, 1);
    setIsUndoDisabled(false);
  };

  const handleCellDoubleClick = (rowIndex, columnKey, value) => {
    console.log('Cell double-clicked:', { rowIndex, columnKey, value });
    setEditCell({ rowIndex, columnKey, value });
  };

  const handleEditChange = (event) => {
    setEditCell({ ...editCell, value: event.target.value });
  };

  const handleEditSave = async () => {
    console.log('Saving edit:', editCell);
    const updatedData = data.map((row, rowIndex) =>
      rowIndex === editCell.rowIndex ? { ...row, [editCell.columnKey]: editCell.value } : row
    );
    setData(updatedData);
    setFilteredData(updatedData);

    const sheet = workbook.worksheets[selectedSheet];
    const headers = sheet.getRow(1).values.slice(1); // Get headers
    const columnIndex = headers.indexOf(editCell.columnKey) + 1; // Find correct column index
    const row = sheet.getRow(editCell.rowIndex + 2); // Get the correct row
    console.log('Updating sheet:', { rowIndex: editCell.rowIndex + 2, columnIndex, value: editCell.value });
    row.getCell(columnIndex).value = editCell.value; // Correctly update the cell value
    row.commit(); // Ensure changes are committed to the workbook

    setEditCell({ rowIndex: null, columnKey: null, value: '' });
    setIsUndoDisabled(false);
  };

  const handleNewRowSave = async () => {
    console.log('Saving new row:', newRow);
    const headers = workbook.worksheets[selectedSheet].getRow(1).values.slice(1);
    const newRowData = headers.reduce((acc, header) => {
      acc[header] = newRow[header] || '';
      return acc;
    }, {});
    const updatedData = [...data, newRowData];
    setData(updatedData);
    setFilteredData(updatedData);
    const sheet = workbook.worksheets[selectedSheet];
    sheet.addRow(newRowData).commit();
    setNewRow({});
    setIsUndoDisabled(false);
  };

  const saveFile = async () => {
    console.log('Saving file');
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'output.xlsx';
    link.click();
  };

  const handlePushToDB = async () => {
    const filteredData = data.filter(row => row.email && row['Phone Number 1']);
    try {
      console.log('Data to be pushed:', filteredData);
      const response = await axios.post('http://localhost:3001/api/pushToDB', filteredData);
      console.log(response.data);
    } catch (error) {
      console.error('Error pushing data to DB: ', error);
    }
  };

  const handleManualSelect = async () => {
    if (selectedManualRow) {
      await mergeRows(selectedManualRow.row, data[currentRowIndex], workbook.worksheets[targetSheetIndex].getRow(1).values.slice(1), workbook.worksheets[targetSheetIndex]);
      setManualSelectDialogOpen(false);
      setMergeDialogOpen(false);
      setIsUndoDisabled(false);
    }
  };

  const handleMergeAll = async () => {
    const targetSheet = workbook.worksheets[targetSheetIndex];
    if (!targetSheet) return;

    const targetHeaders = targetSheet.getRow(1).values.slice(1);
    const targetFirstNameIndex = targetHeaders.indexOf('firstName') + 1;
    const targetLastNameIndex = targetHeaders.indexOf('lastName') + 1;

    const rowsToDelete = [];

    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
      const selectedRow = data[rowIndex];
      let matchingRow;
      targetSheet.eachRow((row) => {
        const rowFirstName = row.getCell(targetFirstNameIndex).value?.toString().toLowerCase().trim();
        const rowLastName = row.getCell(targetLastNameIndex).value?.toString().toLowerCase().trim();
        const selectedFullNameParts = selectedRow.fullName.toLowerCase().split(' ');
        const selectedFirstName = selectedFullNameParts[0].trim();
        const selectedLastName = selectedFullNameParts[selectedFullNameParts.length - 1].trim();

        if (rowFirstName === selectedFirstName && rowLastName === selectedLastName) {
          matchingRow = row;
        }
      });

      if (matchingRow) {
        await mergeRows(matchingRow, selectedRow, targetHeaders, targetSheet);
        rowsToDelete.push(rowIndex);
      }
    }

    const updatedData = data.filter((_, index) => !rowsToDelete.includes(index));
    setData(updatedData);
    setFilteredData(updatedData);

    loadSheetData(workbook.worksheets[selectedSheet]);
    setIsUndoDisabled(false);
  };

  const handleUndo = async () => {
    console.log('Undoing last action');
    const lastAction = undoStack.pop();
    if (!lastAction) return;

    const sheet = workbook.worksheets[lastAction.sheetIndex];
    switch (lastAction.type) {
      case 'delete':
        const restoredData = [...data];
        restoredData.splice(lastAction.index, 0, lastAction.data);
        setData(restoredData);
        setFilteredData(restoredData);
        sheet.insertRow(lastAction.index + 2, lastAction.data);
        break;
      case 'edit':
        const editedData = [...data];
        editedData[lastAction.rowIndex] = lastAction.data;
        setData(editedData);
        setFilteredData(editedData);
        const row = sheet.getRow(lastAction.rowIndex + 2);
        Object.entries(lastAction.data).forEach(([key, value]) => {
          const columnIndex = sheet.getRow(1).values.indexOf(key) + 1;
          row.getCell(columnIndex).value = value;
        });
        break;
      case 'merge':
        const { originalRowValues, newRowValues } = lastAction;
        const rowToUndo = sheet.getRow(Object.keys(originalRowValues)[0]);
        Object.entries(originalRowValues).forEach(([key, value]) => {
          rowToUndo.getCell(key).value = value;
        });
        const targetData = [...data, newRowValues];
        setData(targetData);
        setFilteredData(targetData);
        break;
      default:
        break;
    }

    setUndoStack([...undoStack]);
    setIsUndoDisabled(true);
  };

  const openLinkInPopup = (url) => {
    if (isElectron() && shell) {
      shell.openExternal(url);
    } else {
      const popup = window.open(url, '_blank', 'toolbar=0,location=0,menubar=0,width=800,height=600');
      if (popup) {
        popup.focus();
      }
    }
  };

  const handleSearch = (e) => {
    const searchValue = e.target.value.toLowerCase();
    setSearchText(searchValue);
    setFilteredData(data.filter(row =>
      Object.values(row).some(value =>
        value?.toString().toLowerCase().includes(searchValue)
      )
    ));
  };

  return (
    <Box sx={{ bgcolor: 'background.default', color: 'text.primary', minHeight: '100vh', p: 2 }} ref={tableRef}>
      <input type="file" onChange={handleFileUpload} />
      {workbook && (
        <>
          <Button onClick={handlePushToDB} variant="contained" sx={{ mt: 2, mr: 2 }}>Push to DB</Button>
          <Button onClick={saveFile} variant="contained" sx={{ mt: 2, mr: 2 }}>Save File</Button>
          <Button onClick={handleMergeAll} variant="contained" sx={{ mt: 2, mr: 2 }}>Merge All</Button>
          <Button onClick={handleUndo} variant="contained" sx={{ mt: 2 }} disabled={isUndoDisabled}>Undo</Button>
          <Tabs value={selectedSheet} onChange={handleSheetChange} variant="scrollable" scrollButtons="auto">
            {workbook.worksheets.map((sheet, index) => (
              <Tab label={sheet.name} key={index} />
            ))}
          </Tabs>
          <Box sx={{ display: 'flex', alignItems: 'center', mt: 2 }}>
            <SearchIcon sx={{ mr: 2 }} />
            <TextField
              placeholder="Search..."
              value={searchText}
              onChange={handleSearch}
            />
          </Box>
          <Paper sx={{ mt: 2, height: '70vh', overflow: 'auto' }}>
            <Table stickyHeader>
              <TableHead>
                <TableRow>
                  {workbook.worksheets[selectedSheet].getRow(1).values.slice(1).map((header, index) => (
                    <TableCell key={index} sx={{ position: 'sticky', top: 0, zIndex: 3, backgroundColor: '#b8860b', color: 'black', fontWeight: 'bold' }}>{header}</TableCell>
                  ))}
                  <TableCell sx={{ position: 'sticky', top: 0, zIndex: 3, backgroundColor: '#b8860b', color: 'black', fontWeight: 'bold' }}>
                    <Tooltip title="Add New Row">
                      <IconButton onClick={handleNewRowSave}>
                        <AddCircleIcon />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {filteredData.map((row, rowIndex) => (
                  <TableRow key={rowIndex}>
                    {Object.entries(row).map(([key, value], columnIndex) => (
                      <TableCell
                        key={columnIndex}
                        sx={
                          columnIndex < 2
                            ? {
                                position: 'sticky',
                                left: columnIndex === 0 ? 0 : 100,
                                zIndex: 2,
                                backgroundColor: '#b8860b',
                                color: 'black',
                                fontWeight: 'bold',
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis'
                              }
                            : {
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                color: typeof value === 'string' && value.startsWith('http') ? 'white' : 'inherit'
                              }
                        }
                        onDoubleClick={() => handleCellDoubleClick(rowIndex, key, value)}
                      >
                        {editCell.rowIndex === rowIndex && editCell.columnKey === key ? (
                          <TextField
                            value={editCell.value}
                            onChange={handleEditChange}
                            onBlur={handleEditSave}
                            autoFocus
                          />
                        ) : typeof value === 'string' && value.startsWith('http') ? (
                          <a
                            href="#"
                            onClick={() => openLinkInPopup(value)}
                            style={{ color: 'white', textDecoration: 'underline' }}
                          >
                            {value}
                          </a>
                        ) : (
                          value
                        )}
                      </TableCell>
                    ))}
                    <TableCell
                      sx={{
                        position: 'sticky',
                        right: 0,
                        zIndex: 3,
                        backgroundColor: '#b8860b',
                        color: 'black',
                        fontWeight: 'bold',
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis'
                      }}
                    >
                      <Tooltip title="Delete">
                        <IconButton onClick={() => handleDeleteRow(rowIndex)}>
                          <DeleteIcon />
                        </IconButton>
                      </Tooltip>
                      <Tooltip title="Merge">
                        <IconButton onClick={() => handleMergeClick(rowIndex)}>
                          <MergeTypeIcon />
                        </IconButton>
                      </Tooltip>
                    </TableCell>
                  </TableRow>
                ))}
                <TableRow>
                  {workbook.worksheets[selectedSheet].getRow(1).values.slice(1).map((header, index) => (
                    <TableCell key={index}>
                      <TextField
                        placeholder={`Enter ${header}`}
                        value={newRow[header] || ''}
                        onChange={(e) => setNewRow({ ...newRow, [header]: e.target.value })}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleNewRowSave();
                        }}
                      />
                    </TableCell>
                  ))}
                  <TableCell>
                    <Tooltip title="Add New Row">
                      <IconButton onClick={handleNewRowSave}>
                        <AddCircleIcon />
                      </IconButton>
                    </Tooltip>
                  </TableCell>
                </TableRow>
              </TableBody>
            </Table>
          </Paper>
        </>
      )}

      <Dialog open={mergeDialogOpen} onClose={() => setMergeDialogOpen(false)}>
        <DialogTitle>Select Target Sheet</DialogTitle>
        <DialogContent>
          <Select
            value={targetSheetIndex !== null ? targetSheetIndex : ''}
            onChange={(e) => setTargetSheetIndex(e.target.value)}
            fullWidth
          >
            {workbook && workbook.worksheets.map((sheet, index) => (
              <MenuItem value={index} key={index}>{sheet.name}</MenuItem>
            ))}
          </Select>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setMergeDialogOpen(false)}>Cancel</Button>
          <Button onClick={handleMerge}>Merge</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={manualSelectDialogOpen} onClose={() => setManualSelectDialogOpen(false)}>
        <DialogTitle>Manual Row Selection</DialogTitle>
        <DialogContent>
          <Select
            value={selectedManualRow !== null ? selectedManualRow : ''}
            onChange={(e) => setSelectedManualRow(e.target.value)}
            fullWidth
          >
            {manualSelectOptions.map((option, index) => (
              <MenuItem value={option} key={index}>{`${option.firstName} ${option.lastName}`}</MenuItem>
            ))}
          </Select>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setManualSelectDialogOpen(false)}>Cancel</Button>
          <Button onClick={handleManualSelect}>Select</Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
};

export default ExcelHandler;
