package org.example.data;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DepositDataBuilder implements ExcelExportable{

    private BigDecimal requestId;
    private String firstName;
    private String lastName;
    private String accountNumber;

    @Override
    public Object getValueByColumnIndex(int index) {
        switch (index) {
            case 1:
                return accountNumber;
            case 2:
                return firstName;
            case 3:
                return lastName;
            case 4:
                return requestId;
            default:
                return null;
        }
    }
}
