#include "BlockedServoMotor.h"

#define SERVO_PIN 2
#define SENSOR_PIN 54

#define BUILTIN_LED 13


void setup() {
    Serial.begin(115200);

    BlockedServoMotor my_servo(SERVO_PIN, SENSOR_PIN);

    Serial.println("\nMoving to 0 degrees");
    my_servo.setTarget(0);
    delay(1000);
    Serial.print("Current Position: ");
    Serial.println(my_servo.currentPosition);

    Serial.println("\nMoving to 180 degrees");
    my_servo.setTarget(180);
    delay(1000);
    Serial.print("Current Position: ");
    Serial.println(my_servo.currentPosition);
}

void loop() {

}