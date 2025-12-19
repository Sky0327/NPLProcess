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
      default: '#f5f6f8',
      paper: '#ffffff',
      sidebar: '#1a1a2e',
    },
    text: {
      primary: '#212529',
      secondary: '#6c757d',
    },
    // Status colors
    status: {
      pending: '#9e9e9e',
      inProgress: '#ff9800',
      completed: '#4caf50',
      error: '#f44336',
    },
    // Phase colors
    phase: {
      1: '#aa3142',
      2: '#2196f3',
      3: '#9c27b0',
      4: '#ff9800',
      5: '#4caf50',
    },
    success: {
      main: '#4caf50',
      light: '#81c784',
      dark: '#388e3c',
    },
    warning: {
      main: '#ff9800',
      light: '#ffb74d',
      dark: '#f57c00',
    },
    error: {
      main: '#f44336',
      light: '#e57373',
      dark: '#d32f2f',
    },
    info: {
      main: '#2196f3',
      light: '#64b5f6',
      dark: '#1976d2',
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
      fontSize: '1.75rem',
    },
    h5: {
      fontWeight: 600,
      fontSize: '1.25rem',
    },
    h6: {
      fontWeight: 600,
      fontSize: '1rem',
    },
    subtitle1: {
      fontWeight: 500,
      fontSize: '0.9rem',
    },
    subtitle2: {
      fontWeight: 500,
      fontSize: '0.8rem',
      color: '#6c757d',
    },
    body2: {
      fontSize: '0.85rem',
    },
    caption: {
      fontSize: '0.75rem',
      color: '#6c757d',
    },
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          borderRadius: 8,
          padding: '8px 20px',
          fontWeight: 500,
        },
        sizeSmall: {
          padding: '4px 12px',
          fontSize: '0.8rem',
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          borderRadius: 12,
          boxShadow: '0 1px 3px rgba(0,0,0,0.08)',
          border: '1px solid #e9ecef',
          '&:hover': {
            boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
          },
        },
      },
      variants: [
        {
          props: { variant: 'dashboard' },
          style: {
            borderRadius: 8,
            padding: 0,
          },
        },
        {
          props: { variant: 'compact' },
          style: {
            borderRadius: 8,
            boxShadow: 'none',
            border: '1px solid #dee2e6',
          },
        },
      ],
    },
    MuiCardContent: {
      styleOverrides: {
        root: {
          padding: 16,
          '&:last-child': {
            paddingBottom: 16,
          },
        },
      },
    },
    MuiChip: {
      styleOverrides: {
        root: {
          fontWeight: 500,
          fontSize: '0.75rem',
        },
        sizeSmall: {
          height: 22,
          fontSize: '0.7rem',
        },
      },
    },
    MuiLinearProgress: {
      styleOverrides: {
        root: {
          borderRadius: 4,
          height: 8,
          backgroundColor: '#e9ecef',
        },
        bar: {
          borderRadius: 4,
        },
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          backgroundImage: 'none',
        },
      },
    },
    MuiDrawer: {
      styleOverrides: {
        paper: {
          borderRight: 'none',
        },
      },
    },
    MuiListItemButton: {
      styleOverrides: {
        root: {
          borderRadius: 8,
          marginBottom: 4,
          '&.Mui-selected': {
            backgroundColor: 'rgba(170, 49, 66, 0.08)',
            '&:hover': {
              backgroundColor: 'rgba(170, 49, 66, 0.12)',
            },
          },
        },
      },
    },
    MuiTextField: {
      styleOverrides: {
        root: {
          '& .MuiOutlinedInput-root': {
            borderRadius: 8,
          },
        },
      },
    },
    MuiAlert: {
      styleOverrides: {
        root: {
          borderRadius: 8,
        },
      },
    },
  },
})

export default theme
