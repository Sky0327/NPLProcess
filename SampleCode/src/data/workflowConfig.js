export const PHASES = {
  1: {
    id: 1,
    name: '초기화',
    fullName: 'Phase 1: 초기화 및 설정',
    weight: 10,
    color: '#aa3142',
    tasks: [
      { id: 'init', name: '파일 초기화', type: 'setup' },
      { id: 'reportName', name: '보고서명 설정', type: 'input' },
      { id: 'apiSetup', name: 'ID/API 설정', type: 'input' },
      { id: 'dataLoad', name: '데이터 로드', type: 'action' },
    ],
  },
  2: {
    id: 2,
    name: '데이터 조회',
    fullName: 'Phase 2: 병렬 데이터 조회',
    weight: 25,
    color: '#2196f3',
    parallel: true,
    tasks: [
      { id: 'registry', name: '등기조회', api: 'getRegistryInfo' },
      { id: 'landPrice', name: '공시지가조회', api: 'getLandPrice' },
      { id: 'auction', name: '법원경매조회', api: 'getCourtAuction' },
      { id: 'infocareStats', name: '인포케어 통계', api: 'getInfocareStats' },
      {
        id: 'infocareIntegrated',
        name: '인포케어 통합',
        api: 'getInfocareIntegrated',
      },
      {
        id: 'realEstatePrice',
        name: '실거래가_국토',
        api: 'getRealEstatePrice',
      },
      { id: 'valuemapPrice', name: '실거래가_밸류맵', api: 'getValuemapPrice' },
    ],
  },
  3: {
    id: 3,
    name: '중간 처리',
    fullName: 'Phase 3: 중간 처리',
    weight: 20,
    color: '#9c27b0',
    parallel: false,
    tasks: [
      {
        id: 'kbPrice',
        name: 'KB시세조회',
        api: 'getKBPrice',
        deps: ['registry', 'landPrice'],
      },
      {
        id: 'infocareCases',
        name: '인포케어 사례',
        api: 'getInfocareCases',
        deps: ['infocareIntegrated'],
      },
      {
        id: 'distanceGookto',
        name: '거리계산_국토',
        api: 'calculateDistance',
        deps: ['realEstatePrice'],
      },
      {
        id: 'distanceValuemap',
        name: '거리계산_밸류맵',
        api: 'calculateDistanceValuemap',
        deps: ['valuemapPrice'],
      },
    ],
  },
  4: {
    id: 4,
    name: '리포트 생성',
    fullName: 'Phase 4: 리포트 생성',
    weight: 30,
    color: '#ff9800',
    tasks: [
      {
        id: 'report-담보물정보',
        name: '[2-1] 담보물정보',
        deps: ['registry', 'kbPrice'],
      },
      { id: 'report-감정평가', name: '[2-2] 감정평가', deps: ['kbPrice'] },
      { id: 'report-경매정보', name: '[3] 경매정보', deps: ['auction'] },
      {
        id: 'report-낙찰통계',
        name: '[5-1] 낙찰통계',
        deps: ['infocareStats'],
      },
      {
        id: 'report-낙찰사례',
        name: '[5-2] 낙찰사례',
        deps: ['infocareCases'],
      },
      {
        id: 'report-실거래_국토',
        name: '[6-1] 실거래사례_국토',
        deps: ['distanceGookto'],
      },
      {
        id: 'report-실거래_밸류맵',
        name: '[6-1] 실거래사례_밸류맵',
        deps: ['distanceValuemap'],
      },
    ],
  },
  5: {
    id: 5,
    name: '최종 처리',
    fullName: 'Phase 5: 최종 처리',
    weight: 15,
    color: '#4caf50',
    tasks: [
      { id: 'report-물건지', name: '[0] 물건지', type: 'report' },
      { id: 'report-채권현황', name: '[1] 채권현황', type: 'report' },
      { id: 'pdfConversion', name: '등본PDF 변환', type: 'action' },
      { id: 'xlsxExport', name: 'XLSX 내보내기', type: 'export' },
    ],
  },
}

export const PHASE_ORDER = [1, 2, 3, 4, 5]

export const STATUS = {
  PENDING: 'pending',
  IN_PROGRESS: 'in_progress',
  COMPLETED: 'completed',
  ERROR: 'error',
}

export const LOG_TYPES = {
  INFO: 'info',
  SUCCESS: 'success',
  ERROR: 'error',
  WARNING: 'warning',
}

export const getPhaseById = (id) => PHASES[id]

export const getTaskById = (taskId) => {
  for (const phase of Object.values(PHASES)) {
    const task = phase.tasks.find((t) => t.id === taskId)
    if (task) return { ...task, phaseId: phase.id }
  }
  return null
}

export const getTasksByPhase = (phaseId) => PHASES[phaseId]?.tasks || []

export const getDependencies = (taskId) => {
  const task = getTaskById(taskId)
  return task?.deps || []
}
