import React from 'react'
import { useNavigate, useLocation } from 'react-router-dom'
import {
  AppBar,
  Toolbar,
  Typography,
  Box,
  Tabs,
  Tab,
  Container,
} from '@mui/material'
import { Assessment, Description, Home } from '@mui/icons-material'

const Layout = ({ children }) => {
  const navigate = useNavigate()
  const location = useLocation()

  const getTabValue = () => {
    if (location.pathname === '/') return 0
    if (location.pathname === '/smart-npl1') return 1
    if (location.pathname === '/smart-npl2') return 2
    return 0
  }

  const handleTabChange = (event, newValue) => {
    switch (newValue) {
      case 0:
        navigate('/')
        break
      case 1:
        navigate('/smart-npl1')
        break
      case 2:
        navigate('/smart-npl2')
        break
      default:
        break
    }
  }

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', minHeight: '100vh' }}>
      <AppBar position="static" elevation={2}>
        <Toolbar>
          <Typography variant="h5" component="div" sx={{ flexGrow: 0, mr: 4, fontWeight: 700 }}>
            Samil
          </Typography>
          <Typography variant="h6" component="div" sx={{ flexGrow: 0, fontWeight: 500 }}>
            NPL 평가 시스템
          </Typography>
        </Toolbar>
        <Tabs
          value={getTabValue()}
          onChange={handleTabChange}
          sx={{
            borderBottom: 1,
            borderColor: 'divider',
            '& .MuiTab-root': {
              minHeight: 64,
            },
          }}
        >
          <Tab icon={<Home />} iconPosition="start" label="홈" />
          <Tab icon={<Assessment />} iconPosition="start" label="Smart_NPL1 (데이터 수집)" />
          <Tab icon={<Description />} iconPosition="start" label="Smart_NPL2 (리포트 생성)" />
        </Tabs>
      </AppBar>
      <Container maxWidth="xl" sx={{ flex: 1, py: 4 }}>
        {children}
      </Container>
    </Box>
  )
}

export default Layout

