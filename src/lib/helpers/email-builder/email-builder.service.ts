import { Injectable } from '@nestjs/common';
import { promises as fs } from 'fs';
import { templatesPath } from '../../../share'; 

@Injectable()
export class EmailBuilderService {
  async fillConfirmationTemplate(values: { ACCEPT_LINK: string }) {
    return this.fillEmailTemplate(templatesPath.confirmation, values);
  }
  async fillPasswordResetTemplate(values: { RESET_LINK: string }) {
    return this.fillEmailTemplate(templatesPath.passwordReset, values);
  }
  async fillEmailTemplate(
    templatePath: string,
    replaceValues: object,
  ): Promise<string> {
    let fileContent = await fs.readFile(templatePath, 'utf8');
    for (const key of Object.keys(replaceValues)) {
      fileContent = fileContent.replace(key, replaceValues[key]);
    }
    return fileContent;
  }
}
