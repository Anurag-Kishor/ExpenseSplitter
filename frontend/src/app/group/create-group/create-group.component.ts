import { Component, EventEmitter, Output, inject } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import { GroupService } from '../group.service';
import { Group } from '../group';
import { v4 as uuidv4 } from 'uuid';

@Component({
  selector: 'app-create-group',
  templateUrl: './create-group.component.html',
  styleUrl: './create-group.component.scss'
})
export class CreateGroupComponent {

  groupService: GroupService = inject(GroupService);
  @Output() closeForm = new EventEmitter<boolean>();

  applyForm = new FormGroup({
    groupName: new FormControl('')
  });


  submitApplication() {
    let temp: Group = {
      'id': uuidv4(),
      'groupName': this.applyForm.value.groupName ?? ''
    };

    if (temp.groupName != ''){
      this.groupService.createGroup(temp);
    }
    this.closeForm.emit(false);
  }
}
