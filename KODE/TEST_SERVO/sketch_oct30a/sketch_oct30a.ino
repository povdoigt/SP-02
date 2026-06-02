#include <SCServo.h>
#include <SoftwareSerial.h>
#include "Wire.h"

SMS_STS sms_sts;
SoftwareSerial mySerial(2, 3); // RX, TX

int LEDpin = 13;
void setup()
{
  pinMode(LEDpin,OUTPUT);
  digitalWrite(LEDpin, HIGH);
  mySerial.begin(38400);
  Serial.begin(9600);
  Serial.println("SMS_STS Ping Test");
  sms_sts.pSerial = &mySerial;
  delay(1000);
}

void loop()
{
  int ID = sms_sts.Ping(1);
  if(!sms_sts.getErr()){
    digitalWrite(LEDpin, LOW);
    Serial.print("Servo ID:");
    Serial.println(ID, DEC);
    delay(100);
  }else{
    Serial.println("Ping servo ID error!");
    digitalWrite(LEDpin, HIGH);
    delay(2000);
  }
}
