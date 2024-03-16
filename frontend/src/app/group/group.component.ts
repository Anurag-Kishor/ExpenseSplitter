import { AfterViewInit, Component, ViewChild } from '@angular/core';
import { ListGroupComponent } from './list-group/list-group.component';

@Component({
  selector: 'app-group',
  templateUrl: './group.component.html',
  styleUrl: './group.component.scss'
})
export class GroupComponent implements AfterViewInit {
  toggle = false
  @ViewChild('listGroup', {static: false}) listGroup!: ListGroupComponent;

  ngAfterViewInit() {
    this.listGroup.getAllGroups()
  }

  createGroup() {
    this.toggle = true
    this.listGroup.getAllGroups() 
  }

  closeForm(value: boolean) {
    this.toggle = value
  }
  
}
