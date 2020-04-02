import { Injectable } from '@angular/core';
import { NgbDatepickerI18n, NgbDateStruct } from '@ng-bootstrap/ng-bootstrap';

@Injectable()
export class EsDatepickerService extends NgbDatepickerI18n {

    public getWeekdayShortName(weekday: number): string {
        switch (weekday - 1) {
            case 0:
                return "Lu";
            case 1:
                return "Ma";
            case 2:
                return "Mi";
            case 3:
                return "Ju";
            case 4:
                return "Vi";
            case 5:
                return "Sa";
            case 6:
                return "Do"
            default:
                return "";
        }
    }

    public getMonthShortName(month: number): string {
        return this.getMonthFullName(month).substr(0, 3);
    }

    public getMonthFullName(month: number): string {
        switch (month - 1) {
            case 0:
                return "Enero";
            case 1:
                return "Febrero";
            case 2:
                return "Marzo";
            case 3:
                return "Abril";
            case 4:
                return "Mayo";
            case 5:
                return "Junio";
            case 6:
                return "Julio";
            case 7:
                return "Agosto";
            case 8:
                return "Septiembre";
            case 9:
                return "Octubre";
            case 10:
                return "Noviembre";
            case 11:
                return "Diciembre";
            default:
                return "";
        }
    }

    public getDayAriaLabel(date: NgbDateStruct): string {
        return `${date.day}-${date.month}-${date.year}`;
    }
}
