import React from 'react';
import { CssBaseline, ThemeProvider, createTheme } from '@mui/material';
import ExcelHandler from './ExcelHandler';

const darkTheme = createTheme({
  palette: {
    mode: 'dark',
    background: {
      default: '#121212',
      paper: '#1D1D1D',
    },
    text: {
      primary: '#FFFFFF',
    },
  },
  typography: {
    fontFamily: '"Courier New", Courier, monospace',
  },
});

function App() {
  return (
    <ThemeProvider theme={darkTheme}>
      <CssBaseline />
      <div className="App">
        <header className="App-header">
          <h1>KleanData</h1>
          <ExcelHandler />
        </header>
      </div>
    </ThemeProvider>
  );
}

export default App;
