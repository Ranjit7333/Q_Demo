import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ExcelCompComponent } from './excel-comp/excel-comp.component';

const routes: Routes = [
   //{ path: '', pathMatch: 'full', redirectTo: 'excel' },
 // { path: '', component: ExcelCompComponent }
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
