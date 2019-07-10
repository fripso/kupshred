import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ShredderComponent } from './shredder.component';

describe('ShredderComponent', () => {
  let component: ShredderComponent;
  let fixture: ComponentFixture<ShredderComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ShredderComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ShredderComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
