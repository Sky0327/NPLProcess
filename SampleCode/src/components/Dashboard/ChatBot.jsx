import React, { useState, useRef, useEffect } from 'react'
import {
  Box,
  Typography,
  TextField,
  IconButton,
  Avatar,
  Paper,
  CircularProgress,
} from '@mui/material'
import {
  Send,
  SmartToy,
  Person,
} from '@mui/icons-material'
import useWorkflowStore from '../../store/workflowStore'

const ChatBot = ({ open, width }) => {
  const [messages, setMessages] = useState([
    {
      id: 1,
      type: 'bot',
      text: '안녕하세요! NPL 평가 AI 어시스턴트입니다. 현재 입력된 데이터에 대해 궁금한 점을 질문해 주세요.',
      timestamp: new Date(),
    },
  ])
  const [inputValue, setInputValue] = useState('')
  const [isLoading, setIsLoading] = useState(false)
  const messagesEndRef = useRef(null)

  const {
    projectConfig,
    coverData,
    debtInfo,
    collateralInfo,
  } = useWorkflowStore()

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }

  useEffect(() => {
    scrollToBottom()
  }, [messages])

  const handleSend = async () => {
    if (!inputValue.trim() || isLoading) return

    const userMessage = {
      id: Date.now(),
      type: 'user',
      text: inputValue.trim(),
      timestamp: new Date(),
    }

    setMessages((prev) => [...prev, userMessage])
    setInputValue('')
    setIsLoading(true)

    // Mock AI response with delay
    setTimeout(() => {
      const botResponse = {
        id: Date.now() + 1,
        type: 'bot',
        text: '개발중입니다. AI가 현재까지 입력된 데이터를 바탕으로 응답할 예정입니다.',
        timestamp: new Date(),
      }
      setMessages((prev) => [...prev, botResponse])
      setIsLoading(false)
    }, 1000)
  }

  const handleKeyPress = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleSend()
    }
  }

  const formatTime = (date) => {
    return date.toLocaleTimeString('ko-KR', {
      hour: '2-digit',
      minute: '2-digit',
    })
  }

  if (!open) return null

  return (
    <Box
      sx={{
        width: width,
        height: '100vh',
        display: 'flex',
        flexDirection: 'column',
        position: 'fixed',
        right: 0,
        top: 0,
        bgcolor: 'background.paper',
      }}
    >
      {/* Header */}
      <Box
        sx={{
          p: 2,
          borderBottom: '1px solid',
          borderColor: 'divider',
          display: 'flex',
          alignItems: 'center',
          gap: 1.5,
          bgcolor: 'primary.main',
          color: 'white',
        }}
      >
        <SmartToy />
        <Box>
          <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
            AI 어시스턴트
          </Typography>
          <Typography variant="caption" sx={{ opacity: 0.8 }}>
            NPL 평가 데이터 기반 응답
          </Typography>
        </Box>
      </Box>

      {/* Messages Area */}
      <Box
        sx={{
          flexGrow: 1,
          overflow: 'auto',
          p: 2,
          display: 'flex',
          flexDirection: 'column',
          gap: 2,
          bgcolor: 'grey.50',
          '&::-webkit-scrollbar': {
            width: 6,
          },
          '&::-webkit-scrollbar-thumb': {
            bgcolor: 'grey.300',
            borderRadius: 3,
          },
        }}
      >
        {messages.map((message) => (
          <Box
            key={message.id}
            sx={{
              display: 'flex',
              flexDirection: message.type === 'user' ? 'row-reverse' : 'row',
              alignItems: 'flex-start',
              gap: 1,
            }}
          >
            <Avatar
              sx={{
                width: 32,
                height: 32,
                bgcolor: message.type === 'user' ? 'secondary.main' : 'primary.main',
              }}
            >
              {message.type === 'user' ? (
                <Person sx={{ fontSize: 18 }} />
              ) : (
                <SmartToy sx={{ fontSize: 18 }} />
              )}
            </Avatar>
            <Box
              sx={{
                maxWidth: '80%',
              }}
            >
              <Paper
                elevation={0}
                sx={{
                  p: 1.5,
                  bgcolor: message.type === 'user' ? 'primary.main' : 'white',
                  color: message.type === 'user' ? 'white' : 'text.primary',
                  borderRadius: 2,
                  borderTopRightRadius: message.type === 'user' ? 0 : 2,
                  borderTopLeftRadius: message.type === 'user' ? 2 : 0,
                }}
              >
                <Typography variant="body2" sx={{ whiteSpace: 'pre-wrap' }}>
                  {message.text}
                </Typography>
              </Paper>
              <Typography
                variant="caption"
                sx={{
                  color: 'text.secondary',
                  display: 'block',
                  mt: 0.5,
                  textAlign: message.type === 'user' ? 'right' : 'left',
                  fontSize: '0.65rem',
                }}
              >
                {formatTime(message.timestamp)}
              </Typography>
            </Box>
          </Box>
        ))}

        {isLoading && (
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
            <Avatar
              sx={{
                width: 32,
                height: 32,
                bgcolor: 'primary.main',
              }}
            >
              <SmartToy sx={{ fontSize: 18 }} />
            </Avatar>
            <Paper
              elevation={0}
              sx={{
                p: 1.5,
                bgcolor: 'white',
                borderRadius: 2,
                borderTopLeftRadius: 0,
                display: 'flex',
                alignItems: 'center',
                gap: 1,
              }}
            >
              <CircularProgress size={16} />
              <Typography variant="body2" color="text.secondary">
                응답 생성 중...
              </Typography>
            </Paper>
          </Box>
        )}

        <div ref={messagesEndRef} />
      </Box>

      {/* Input Area */}
      <Box
        sx={{
          p: 2,
          borderTop: '1px solid',
          borderColor: 'divider',
          bgcolor: 'background.paper',
        }}
      >
        <Box sx={{ display: 'flex', gap: 1 }}>
          <TextField
            fullWidth
            size="small"
            placeholder="메시지를 입력하세요..."
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            onKeyPress={handleKeyPress}
            disabled={isLoading}
            sx={{
              '& .MuiOutlinedInput-root': {
                borderRadius: 3,
              },
            }}
          />
          <IconButton
            color="primary"
            onClick={handleSend}
            disabled={!inputValue.trim() || isLoading}
            sx={{
              bgcolor: 'primary.main',
              color: 'white',
              '&:hover': {
                bgcolor: 'primary.dark',
              },
              '&.Mui-disabled': {
                bgcolor: 'grey.300',
                color: 'grey.500',
              },
            }}
          >
            <Send />
          </IconButton>
        </Box>
        <Typography
          variant="caption"
          sx={{ color: 'text.secondary', display: 'block', mt: 1, textAlign: 'center' }}
        >
          현재 입력된 데이터: 채권 {debtInfo.length}건, 담보물 {collateralInfo.length}건
        </Typography>
      </Box>
    </Box>
  )
}

export default ChatBot
