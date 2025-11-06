#include <Servo.h>
#include "SmallMovingAverage.h"


BlockedServoMotor::BlockedServoMotor(int servoPin, int sensorPin) : filter(10) {
    init(servoPin, sensorPin);
}

void BlockedServoMotor::init(int servoPin, int sensorPin) {
    this->sensorPin = sensorPin;
    servo.attach(servoPin);
    pinMode(this->sensorPin, INPUT);
    this->currentPosition = 90;
    filter.init();
}

int BlockedServoMotor::setTarget(int target) {

    int direction = (target > this->currentPosition) ? 1 : -1;

    servo.write(target);

    delay(500); // Wait for servo to reach position

    int val = 0;
    int blocked = 0;

    // Simple loop to monitor the sensor value
    while (true) {
        for (int i = 0; i < 10; i++) {
            val = this->filter.update(analogRead(this->sensorPin));
        }
        if (val >= 1020) {
            break; // Exit loop if sensor value indicates movement
        }
        blocked = 1;
        // Servo is blocked
        target -= direction;
        servo.write(target);
        delay(10); // Wait for servo to attempt to reach position
    }
    this->currentPosition = target;
    return blocked;
}
