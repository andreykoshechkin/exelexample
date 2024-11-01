package org.example.data;

import lombok.*;
import org.example.annotion.ExcelPresenter;

import java.time.LocalDate;

@Builder
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class CreditDataBuilder extends ExcelDataModel {

    @ExcelPresenter(name = "Дата взятие кредита")
    private LocalDate dateCreditOpen;
    @ExcelPresenter(name = "Имя")
    private String firstName;
    @ExcelPresenter(name = "Фамилия")
    private String lastName;
    @ExcelPresenter(name = "Первоночальный взнос")
    private String firstAmount;
    @ExcelPresenter(name = "Процент %")
    private String procent;
}
