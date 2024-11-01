package org.example.data;


import lombok.*;
import org.example.annotion.ExcelPresenter;

import java.math.BigDecimal;


@Builder
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class DepositDataBuilder extends ExcelDataModel {

    @ExcelPresenter(name = "Номер заявки")
    private BigDecimal requestId;
    @ExcelPresenter(name = "Имя")
    private String firstName;
    @ExcelPresenter(name = "Фамилия")
    private String lastName;
    @ExcelPresenter(name = "Депозитный счет")
    private String accountNumber;

}
