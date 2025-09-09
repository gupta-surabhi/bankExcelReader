public class TicketAgeBoundDetails {
    private int lowerBound;
    private int upperBound;
    private String columnName;
    private int count;

    public TicketAgeBoundDetails(int lowerBound, int upperBound, String columnName) {
        this.lowerBound = lowerBound;
        this.upperBound = upperBound;
        this.columnName = columnName;
        this.count = 0;
    }

    public int getLowerBound() {
        return lowerBound;
    }

    public void setLowerBound(int lowerBound) {
        this.lowerBound = lowerBound;
    }

    public int getUpperBound() {
        return upperBound;
    }

    public void setUpperBound(int upperBound) {
        this.upperBound = upperBound;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public void incrementCount() {
        this.count = this.count+1;
    }

    public void checkAndIncrementCount(int age) {
        if(age >= this.lowerBound && age <= this.upperBound)
            incrementCount();
    }
}
