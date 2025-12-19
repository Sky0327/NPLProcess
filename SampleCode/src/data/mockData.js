// Mock 데이터 생성

export const mockRegistryData = [
  {
    id: 1,
    registryNumber: '12345-01-123456',
    address: '서울특별시 강남구 테헤란로 123',
    owner: '홍길동',
    area: '85.5㎡',
    usage: '상업지역',
    rights: '소유권',
  },
  {
    id: 2,
    registryNumber: '12345-02-234567',
    address: '서울특별시 서초구 서초대로 456',
    owner: '김철수',
    area: '120.3㎡',
    usage: '주거지역',
    rights: '소유권',
  },
]

export const mockLandPriceData = [
  {
    id: 1,
    registryNumber: '12345-01-123456',
    address: '서울특별시 강남구 테헤란로 123',
    landType: '상업지역',
    area: '85.5㎡',
    officialPrice: '1,500,000,000원',
    year: '2024',
  },
  {
    id: 2,
    registryNumber: '12345-02-234567',
    address: '서울특별시 서초구 서초대로 456',
    landType: '주거지역',
    area: '120.3㎡',
    officialPrice: '2,300,000,000원',
    year: '2024',
  },
]

export const mockAuctionData = [
  {
    id: 1,
    registryNumber: '12345-01-123456',
    court: '서울중앙지방법원',
    caseNumber: '2024타경12345',
    startDate: '2024-01-15',
    endDate: '2024-02-15',
    price: '1,200,000,000원',
    status: '진행중',
  },
]

export const mockKBData = [
  {
    id: 1,
    registryNumber: '12345-01-123456',
    address: '서울특별시 강남구 테헤란로 123',
    kbPrice: '1,450,000,000원',
    area: '85.5㎡',
    usage: '상업',
    buildYear: '2010',
  },
]

export const mockRealEstatePriceData = [
  {
    id: 1,
    registryNumber: '12345-01-123456',
    address: '서울특별시 강남구 테헤란로 123',
    transactionDate: '2023-12-15',
    price: '1,380,000,000원',
    area: '85.5㎡',
    distance: '500m',
  },
]

export const mockInfocareData = [
  {
    id: 1,
    registryNumber: '12345-01-123456',
    address: '서울특별시 강남구 테헤란로 123',
    period: '1개월',
    scope: '동',
    count: 5,
    averagePrice: '1,400,000,000원',
  },
]

export const mockReportData = {
  collateralInfo: {
    totalCount: 10,
    totalValue: '15,000,000,000원',
    averageValue: '1,500,000,000원',
  },
  appraisal: {
    totalCount: 8,
    averagePrice: '1,450,000,000원',
  },
  auctionInfo: {
    totalCount: 3,
    averagePrice: '1,200,000,000원',
  },
}

