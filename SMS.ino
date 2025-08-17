#include <Arduino.h>
#define SECOND 1000

#define led1 0
#define PIR_PIN 15

int lastPirVal = LOW;
int pirVal;

unsigned long motionStartTime = 0;
unsigned long motionEndTime = 0;
unsigned long currentTime = 0;
unsigned long idleStartTime = 0;

void setup()
{
    pinMode(led1, OUTPUT);
    pinMode(PIR_PIN, INPUT);
    Serial.begin(115200);
}

void loop()
{
    pirVal = digitalRead(PIR_PIN);
    if (pirVal == 1)
    {
        digitalWrite(led1, HIGH);
        if (lastPirVal == LOW)
        {
            motionStartTime = millis();
            if (idleStartTime != 0)
            {
                unsigned long idleDuration = (millis() - idleStartTime) / SECOND;
                currentTime += idleDuration;
                Serial.print("Current time: ");
                Serial.print(currentTime);
                Serial.print(" seconds Duration: ");
                Serial.print(idleDuration);
                Serial.println(" seconds. (Idle)");
            }
            lastPirVal = HIGH;
        }
    }
    else
    {
        digitalWrite(led1, LOW);
        if (lastPirVal == HIGH)
        {
            unsigned long motionDuration = (millis() - motionStartTime) / SECOND;
            if (motionDuration > 0)
            {
                currentTime += motionDuration;
                Serial.print("Current time: ");
                Serial.print(currentTime);
                Serial.print(" seconds Duration: ");
                Serial.print(motionDuration);
                Serial.println(" seconds. (Movement)");
            }

            idleStartTime = millis();
            lastPirVal = LOW;
        }
        else
        {
            if (idleStartTime == 0)
            {
                idleStartTime = millis();
            }
        }
    }
}
