#ifndef SmallMovingAverage_h
#define SmallMovingAverage_h

class SmallMovingAverage {
  public:
    SmallMovingAverage(int size);
    void init();
    int update(int input);

  private:
    int* values;
    int index;
    int count;
    int sum;
    int size;
};

#endif // SmallMovingAverage_h