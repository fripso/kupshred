import { Component } from '@angular/core';
import { XlsService } from 'src/app/services/xls.service';

@Component({
    selector: 'app-shredder',
    templateUrl: './shredder.component.html',
    styleUrls: ['./shredder.component.scss']
})
export class ShredderComponent {

    values = {
        ak: 35,
        al: 60,
        am: 65
    };

    locale = 'nl-NL';

    constructor(private xls: XlsService) { }

    loadXLS(event) {
        if (event.target.files.length >= 1) {
        this.xls.loadXLS(event.target.files, this.values, this.locale);
        }
    }
}
