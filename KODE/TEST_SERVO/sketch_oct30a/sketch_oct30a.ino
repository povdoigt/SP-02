

void setup() {
  pinMode(A1, OUTPUT);

  Serial.begin();
}

void loop() {
  digitalWrite(A1, HIGH);
  delay(200);
  digitalWrite(A1, LOW);
  delay(200);

  Serial.println("Hello !");
}
