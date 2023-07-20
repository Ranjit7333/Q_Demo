import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { ExcelCompComponent } from './excel-comp/excel-comp.component';
@NgModule({
  declarations: [
    AppComponent,
   
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    ExcelCompComponent
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
