#ifndef __SCS_SERAIL_H__
#define __SCS_SERAIL_H__

void ftUart_Send(uint8_t *nDat, int nLen);
int ftUart_Read(uint8_t *nDat, int nLen);
void ftBus_Delay(void);

//UART 接收数据接口
int readSCS(unsigned char *nDat, int nLen);

//UART 发送数据接口
int writeSCS(unsigned char *nDat, int nLen);

int writeByteSCS(unsigned char bDat);

//接收缓冲区刷新
void rFlushSCS();

//发送缓冲区刷新
void wFlushSCS();

#endif