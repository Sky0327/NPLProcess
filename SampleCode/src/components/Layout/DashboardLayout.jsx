import React, { useState } from 'react'
import {
  Box,
  AppBar,
  Toolbar,
  Typography,
  IconButton,
  useTheme,
  useMediaQuery,
} from '@mui/material'
import {
  Menu as MenuIcon,
  ChevronLeft,
  ChevronRight,
} from '@mui/icons-material'
import GlobalProgressBar from '../Dashboard/GlobalProgressBar'
import WorkflowSidebar from '../Dashboard/WorkflowSidebar'
import ChatBot from '../Dashboard/ChatBot'

const SIDEBAR_WIDTH = 260
const CHATBOT_WIDTH = 360

const DashboardLayout = ({ children }) => {
  const theme = useTheme()
  const isMobile = useMediaQuery(theme.breakpoints.down('md'))
  const [sidebarOpen, setSidebarOpen] = useState(!isMobile)
  const [chatBotOpen, setChatBotOpen] = useState(!isMobile)

  return (
    <Box sx={{ display: 'flex', minHeight: '100vh', bgcolor: 'background.default' }}>
      {/* Sidebar */}
      <Box
        sx={{
          width: sidebarOpen ? SIDEBAR_WIDTH : 0,
          flexShrink: 0,
          transition: 'width 0.2s ease',
          overflow: 'hidden',
        }}
      >
        <WorkflowSidebar open={sidebarOpen} width={SIDEBAR_WIDTH} />
      </Box>

      {/* Main Content Area */}
      <Box
        sx={{
          flexGrow: 1,
          display: 'flex',
          flexDirection: 'column',
          minWidth: 0,
        }}
      >
        {/* Header */}
        <AppBar
          position="sticky"
          elevation={0}
          sx={{
            bgcolor: 'background.paper',
            borderBottom: '1px solid',
            borderColor: 'divider',
          }}
        >
          <Toolbar sx={{ minHeight: 56, px: 2 }}>
            <IconButton
              edge="start"
              onClick={() => setSidebarOpen(!sidebarOpen)}
              sx={{ mr: 2, color: 'text.primary' }}
            >
              {sidebarOpen ? <ChevronLeft /> : <MenuIcon />}
            </IconButton>

            <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5 }}>
              <Box
                component="img"
                src="/samil-logo.png"
                alt="Samil"
                sx={{ height: 28 }}
                onError={(e) => {
                  e.target.style.display = 'none'
                }}
              />
              <Typography
                variant="h6"
                sx={{
                  color: 'primary.main',
                  fontWeight: 700,
                  fontSize: '1.1rem',
                }}
              >
                NPL 평가 시스템
              </Typography>
            </Box>

            <Box sx={{ flexGrow: 1 }} />

            <IconButton
              onClick={() => setChatBotOpen(!chatBotOpen)}
              sx={{ color: 'text.primary' }}
            >
              {chatBotOpen ? <ChevronRight /> : <ChevronLeft />}
            </IconButton>
          </Toolbar>

          {/* Global Progress Bar */}
          <GlobalProgressBar />
        </AppBar>

        {/* Page Content */}
        <Box
          component="main"
          sx={{
            flexGrow: 1,
            p: 3,
            overflow: 'auto',
          }}
        >
          {children}
        </Box>
      </Box>

      {/* ChatBot Sidebar */}
      <Box
        sx={{
          width: chatBotOpen ? CHATBOT_WIDTH : 0,
          flexShrink: 0,
          transition: 'width 0.2s ease',
          overflow: 'hidden',
          borderLeft: chatBotOpen ? '1px solid' : 'none',
          borderColor: 'divider',
          bgcolor: 'background.paper',
        }}
      >
        <ChatBot open={chatBotOpen} width={CHATBOT_WIDTH} />
      </Box>
    </Box>
  )
}

export default DashboardLayout
