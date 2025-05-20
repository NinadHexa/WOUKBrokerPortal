import com.poiji.annotation.*;
public class CalculateModel {
	@ExcelCellName("Total Loan Amount")
    public String loanAmount;

    @ExcelCellName("Advance Payments")
    public String advancePayments;

    @ExcelCellName("Term")
    public String term;

    @ExcelCellName("Risk Band")
    public String riskBand;

    @ExcelCellName("Commission")
    public String commission;

    @ExcelCellName("Expected Commission")
    public String expectedCommission;

    @ExcelCellName("received Commission")
    public String receivedCommission;

    @ExcelCellName("Expected Net Yield")
    public String expectedNetYield;

    @ExcelCellName("Received Net Yield")
    public String receivedNetYield;

    // Getters
    public String getLoanAmount() {
        return loanAmount;
    }

    public String getAdvancePayments() {
        return advancePayments;
    }

    public String getTerm() {
        return term;
    }

    public String getRiskBand() {
        return riskBand;
    }

    public String getCommission() {
        return commission;
    }

    public String getExpectedCommission() {
        return expectedCommission;
    }

    public String getReceivedCommission() {
        return receivedCommission;
    }

    public String getExpectedNetYield() {
        return expectedNetYield;
    }

    public String getReceivedNetYield() {
        return receivedNetYield;
    }

    // Setters
    public void setLoanAmount(String loanAmount) {
        this.loanAmount = loanAmount;
    }

    public void setAdvancePayments(String advancePayments) {
        this.advancePayments = advancePayments;
    }

    public void setTerm(String term) {
        this.term = term;
    }

    public void setRiskBand(String riskBand) {
        this.riskBand = riskBand;
    }

    public void setCommission(String commission) {
        this.commission = commission;
    }

    public void setExpectedCommission(String expectedCommission) {
        this.expectedCommission = expectedCommission;
    }

    public void setReceivedCommission(String receivedCommission) {
        this.receivedCommission = receivedCommission;
    }

    public void setExpectedNetYield(String expectedNetYield) {
        this.expectedNetYield = expectedNetYield;
    }

    public void setReceivedNetYield(String receivedNetYield) {
        this.receivedNetYield = receivedNetYield;
    }

}
