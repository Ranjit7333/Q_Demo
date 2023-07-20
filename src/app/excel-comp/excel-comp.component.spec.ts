import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelCompComponent } from './excel-comp.component';

describe('ExcelCompComponent', () => {
  let component: ExcelCompComponent;
  let fixture: ComponentFixture<ExcelCompComponent>;

  beforeEach(() => {
    TestBed.configureTestingModule({
      imports: [ExcelCompComponent]
    });
    fixture = TestBed.createComponent(ExcelCompComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
