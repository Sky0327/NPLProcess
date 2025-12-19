import { createTheme } from '@mui/material/styles'

// Samil 브랜드 컬러 - 짙은 버건디
const burgundy = {
  50: '#f5e6e8',
  100: '#e6c1c6',
  200: '#d598a0',
  300: '#c46f7a',
  400: '#b7505e',
  500: '#aa3142', // 메인 브랜드 컬러
  600: '#a32c3c',
  700: '#9a2533',
  800: '#911f2b',
  900: '#80131e',
}

const theme = createTheme({
  palette: {
    primary: {
      main: burgundy[500],
      light: burgundy[300],
      dark: burgundy[700],
      contrastText: '#ffffff',
    },
    secondary: {
      main: '#6c757d',
      light: '#adb5bd',
      dark: '#495057',
    },
    background: {
      default: '#f8f9fa',
      paper: '#ffffff',
    },
    text: {
      primary: '#212529',
      secondary: '#6c757d',
    },
  },
  typography: {
    fontFamily: [
      '-apple-system',
      'BlinkMacSystemFont',
      '"Segoe UI"',
      'Roboto',
      '"Helvetica Neue"',
      'Arial',
      'sans-serif',
    ].join(','),
    h4: {
      fontWeight: 600,
    },
    h5: {
      fontWeight: 600,
    },
    h6: {
      fontWeight: 600,
    },
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          borderRadius: 8,
          padding: '8px 24px',
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          borderRadius: 12,
          boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
        },
      },
    },
  },
})

export default theme

