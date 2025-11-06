#ifndef BlockedServoMotor_h
#define BlockedServoMotor_h

#include <Servo.h>
#include "SmallMovingAverage.h"

class BlockedServoMotor {
  public:
    int currentPosition;
    BlockedServoMotor(int servoPin, int sensorPin);
    void init(int servoPin, int sensorPin);
    int setTarget(int target);

  private:
    Servo servo;
    int sensorPin;
    SmallMovingAverage filter;
};

#endif // BlockedServoMotor_h