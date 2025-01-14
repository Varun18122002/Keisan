import { Component } from '@angular/core';
import { HeaderComponent } from "./component/header/header.component";
import { ToolbarComponent } from "./component/toolbar/toolbar.component";
import { GridComponent } from "./component/grid/grid.component";

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [HeaderComponent,   GridComponent],
  template: `
    <div class="app-container fixed">
      <app-header></app-header>
      <div class="grid-container">
        <app-grid></app-grid>
      </div>
    </div>
  `,
  styles: `
  .app-container{
    display: flex;
    flex-direction: column;
    height: 95vh;
    overflow: hidden;
  }

  app-header , app-toolbar{
    flex-shrink: 0;
    z-index: 10;
  }

  .grid-container{
    flex: 1;
    overflow: auto;
  }

  `,
})
export class AppComponent {
  title = 'Keisan';
}
