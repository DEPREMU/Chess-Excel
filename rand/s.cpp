int pinLedGreen = 12;
int pinLedYellow = 8;
int pinLedRed = 7;
int pinButton = 2;
bool turnOn = false;
int count = 1;

void setup() {
    Serial.begin(9600);
    pinMode(pinLedGreen, OUTPUT);
    pinMode(pinLedYellow, OUTPUT);
    pinMode(pinLedRed, OUTPUT);
    pinMode(pinButton, INPUT);
}

void turnOffLeds() {
    digitalWrite(pinLedRed, LOW);
    digitalWrite(pinLedYellow, LOW);
    digitalWrite(pinLedGreen, LOW);
}

bool turnOff() {
    turnOffLeds();
    count = 0;
    Serial.println("Turned off");
    return true;
}

bool waitTimer(int timer) {
    int start = millis();
        delay(75);
        if (digitalRead(pinButton) != HIGH) continue; // Makes a continue if the button is not pressed

        turnOn = !turnOn; // exchange the value, from (true to false) or (false to true)
        if(turnOn) continue;  // Makes a continue if it's not turned off, this means the traffic light will continue
        return turnOff(); // Call the function turnOff
    }
    return false; // Return false, this means that the delay finished and the program wasn't interrupted
}

void loop() {
  waitTimer(200);

  if (!turnOn) return;  // Makes a early return, if turnOn is true it'll continue
  
  count++; // Increment the count
  Serial.println("Number of executes: " + String(count)); // Shows a string with the number of executes
  if (waitTimer(100)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 100 ms

  digitalWrite(pinLedGreen, HIGH); // Turn on the green led
  if (waitTimer(5000)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 5 seconds
  
  digitalWrite(pinLedGreen, LOW); // Turn off the green led
  if (waitTimer(100)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 100 ms

  digitalWrite(pinLedYellow, HIGH); // Turn on the yellow led
  if (waitTimer(500)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 500 ms
  
  digitalWrite(pinLedYellow, LOW); // Turn off the yellow led
  if (waitTimer(250)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 250 ms
  
  digitalWrite(pinLedYellow, HIGH); // Turn on the yellow led
  if (waitTimer(500)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 500 ms
  
  digitalWrite(pinLedYellow, LOW); // Turn off the yellow led
  if (waitTimer(250)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 250 ms
  
  digitalWrite(pinLedYellow, HIGH); // Turn on the yellow led
  if (waitTimer(500)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 500 ms
  
  digitalWrite(pinLedYellow, LOW); // Turn off the yellow led
  if (waitTimer(250)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 250 ms
  
  digitalWrite(pinLedRed, HIGH); // Turn on the red led
  if (waitTimer(5000)) return; // Evaluate if the button is pressed in the funcion waitTimer with a delay of 5 seconds
  
  digitalWrite(pinLedRed, LOW); // Turn off the red led
  
}
