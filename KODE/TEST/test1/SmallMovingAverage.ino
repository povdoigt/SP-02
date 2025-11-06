#include "SmallMovingAverage.h"


SmallMovingAverage::SmallMovingAverage(int size) {
    this->size = size;
    values = new int[size];
    init();
}

void SmallMovingAverage::init() {
    memset(values, 0, sizeof(int) * size);
    index = 0;
    count = 0;
    sum = 0;
}

int SmallMovingAverage::update(int input) {

    // Remove the oldest value from sum
    sum -= values[index];

    // Add the new value to the array and sum
    values[index] = input;
    sum += input;

    // Update index and count
    index = (index + 1) % size;
    if (count < size) {
        count++;
    }

    // Calculate and return the average
    return sum / count;
}
