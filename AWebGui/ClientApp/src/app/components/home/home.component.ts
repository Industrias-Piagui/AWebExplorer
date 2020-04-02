import { Component, OnInit } from '@angular/core';
import { NgbDate, NgbCalendar, NgbDateStruct } from '@ng-bootstrap/ng-bootstrap';
import { HttpClient } from '@angular/common/http';
import * as moment from 'moment';

@Component({
  selector: 'home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {
  public hoveredDate: NgbDate;
  public fromDate: NgbDate;
  public toDate: NgbDate;
  public maxDate: NgbDateStruct;
  public disabledButton = false;
  public showAlert = false;
  public showErroAlert = false;

  public constructor(private http: HttpClient, calendar: NgbCalendar) {
    this.toDate = this.fromDate = calendar.getPrev(calendar.getToday(), 'd', 1);
  }

  public ngOnInit(): void {
    var yesterday = moment().add(-1, 'd');
    this.maxDate = {
      year: yesterday.year(),
      month: (yesterday.month() + 1),
      day: yesterday.date()
    };
  }

  onDateSelection(date: NgbDate) {
    if (!this.fromDate && !this.toDate) {
      this.fromDate = date;
    } else if (this.fromDate && !this.toDate && date.after(this.fromDate)) {
      this.toDate = date;
    } else {
      this.toDate = null;
      this.fromDate = date;
    }
  }

  isHovered(date: NgbDate) {
    return this.fromDate && !this.toDate && this.hoveredDate && date.after(this.fromDate) && date.before(this.hoveredDate);
  }

  isInside(date: NgbDate) {
    return date.after(this.fromDate) && date.before(this.toDate);
  }

  isRange(date: NgbDate) {
    return date.equals(this.fromDate) || date.equals(this.toDate) || this.isInside(date) || this.isHovered(date);
  }

  public runRobot(event: Event): void {
    event.preventDefault();
    let from = this.fromDate;
    let to = this.toDate;
    this.disabledButton = true;
    this.http.post('/api/Sales', {
      from: new Date(from.year, from.month - 1, from.day),
      to: new Date(to.year, to.month - 1, to.day),
    }).toPromise().then(res => {
      this.disabledButton = false;
      this.showAlert = true;
    }).catch(err => {
      this.disabledButton = false;
      this.showErroAlert = true;
      console.error(err);
    });
  }
}
