import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { RouterModule } from '@angular/router';
import { HttpClientModule, HTTP_INTERCEPTORS } from '@angular/common/http';
import { NgbModule, NgbDatepickerI18n } from '@ng-bootstrap/ng-bootstrap';

import { AppComponent } from './components/app/app.component';
import { HomeComponent } from './components/home/home.component';

import { EsDatepickerService } from './services/es-datepicker.service';

@NgModule({
  declarations: [
    AppComponent,
    HomeComponent
  ],
  imports: [
    BrowserModule,
    HttpClientModule,
    NgbModule,
    RouterModule.forRoot([
      { path: '', component: HomeComponent, pathMatch: 'full' }
    ]),
  ],
  providers: [
    {
      provide: NgbDatepickerI18n,
      useClass: EsDatepickerService
    }
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
