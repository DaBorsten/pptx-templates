import { log } from '../helper/logger';
import { AutomizerError } from '../errors';
import { FileHelper } from '../helper/file-helper';
import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { Shape } from '../classes/shape';
import { ImportedElement, ShapeTargetType, Target } from '../types/types';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { IShapeAction } from '../interfaces/ishape-action';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import path from 'path';

export class OLEObject extends Shape implements IShapeAction {
  /** The copied embedding, e.g. ppt/embeddings/oleObject2.bin. */
  private embeddingFile: string;
  private embeddingContentType: string;

  constructor(
    shape: ImportedElement,
    targetType: ShapeTargetType,
    sourceArchive: IArchive,
  ) {
    super(shape, targetType);
    this.sourceArchive = sourceArchive;
    this.relRootTag = 'p:oleObj';
    this.relAttribute = 'r:id';
  }

  /**
   * The extension of an embedding, kept as it is: an embedded package
   * (.xlsx, .xlsm, .docx, …) is opened by its extension and content type.
   */
  private getFileExtension(file?: string): string {
    const ext = file ? path.extname(file).toLowerCase() : '';
    return ext || '.bin';
  }

  /**
   * A file name in ppt/embeddings that is not taken yet, after the source
   * name: oleObject1.bin → oleObject2.bin. The name must not depend on the
   * slide's rIds: those repeat from slide to slide, and two embeddings of
   * different slides would overwrite each other.
   */
  private getFreeTargetFile(sourceFilePath: string): string {
    const ext = this.getFileExtension(sourceFilePath);
    const base =
      path.basename(sourceFilePath, path.extname(sourceFilePath)).replace(
        /\d+$/,
        '',
      ) || 'oleObject';
    for (let n = 1; ; n++) {
      const candidate = PptPaths.embedding(`${base}${n}${ext}`);
      if (!this.targetArchive.fileExists(candidate)) return candidate;
    }
  }

  /** The content type of the embedding in the source deck. */
  private async getSourceContentType(sourceFilePath: string): Promise<string> {
    const ext = this.getFileExtension(sourceFilePath);
    const fallback = this.getContentType(ext);
    if (!this.sourceArchive.fileExists('[Content_Types].xml')) return fallback;
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      this.sourceArchive,
      '[Content_Types].xml',
    );
    const override = Array.from(
      contentTypesXml.getElementsByTagName('Override'),
    ).find(
      (element) =>
        element.getAttribute('PartName')?.toLowerCase() ===
        PptPaths.partName(sourceFilePath).toLowerCase(),
    );
    const byDefault = Array.from(
      contentTypesXml.getElementsByTagName('Default'),
    ).find(
      (element) =>
        element.getAttribute('Extension')?.toLowerCase() === ext.slice(1),
    );
    return (
      override?.getAttribute('ContentType') ||
      byDefault?.getAttribute('ContentType') ||
      fallback
    );
  }

  /**
   * Appending OLE objects is not supported; only remove() is implemented.
   */
  async append(): Promise<OLEObject> {
    throw new AutomizerError(
      'OLE objects cannot be appended to a slide; only remove() is supported.',
    );
  }

  /**
   * Modifying OLE objects is not supported; only remove() is implemented.
   */
  async modify(): Promise<OLEObject> {
    throw new AutomizerError(
      'OLE objects cannot be modified; only remove() is supported.',
    );
  }

  // TODO: remove is not currently properly implemented,
  //  suggest we delete the file from the archive as well as removing the relationship.
  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<OLEObject> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();
    await this.removeOleObjectFile();
    await this.removeFromContentTypes();
    await this.removeFromSlideRels();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    oleObjects?: Target[],
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    const allOleObjects =
      oleObjects ||
      (await OLEObject.getAllOnSlide(
        this.sourceArchive,
        this.targetSlideRelFile,
      ));

    const oleObject = allOleObjects.find((obj) => obj.rId === this.sourceRid);
    if (!oleObject) {
      throw new Error(`OLE object with rId ${this.sourceRid} not found.`);
    }

    const sourceFilePath = PptPaths.embedding(oleObject.file.split('/').pop());

    this.createdRid = await XmlRelationshipHelper.getNextRelId(
      this.targetArchive,
      this.targetSlideRelFile,
    );

    this.embeddingFile = this.getFreeTargetFile(sourceFilePath);
    this.embeddingContentType = await this.getSourceContentType(sourceFilePath);

    await this.copyOleObjectFile(sourceFilePath);
    await this.appendToContentTypes();
    await this.updateSlideRels();
    await this.updateSlideXml();
  }

  private async copyOleObjectFile(sourceFilePath: string): Promise<void> {
    try {
      await FileHelper.zipCopy(
        this.sourceArchive,
        sourceFilePath,
        this.targetArchive,
        this.embeddingFile,
      );
    } catch (error) {
      log.error('Error copying OLE object file:', error);
      throw error;
    }
  }

  private async appendToContentTypes(): Promise<void> {
    const contentTypesPath = '[Content_Types].xml';
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      contentTypesPath,
    );

    const types = contentTypesXml.getElementsByTagName('Types')[0];
    const partName = PptPaths.partName(this.embeddingFile);
    const existingOverride = Array.from(
      types.getElementsByTagName('Override'),
    ).find((override) => override.getAttribute('PartName') === partName);

    if (!existingOverride) {
      const newOverride = contentTypesXml.createElement('Override');
      newOverride.setAttribute('PartName', partName);
      newOverride.setAttribute('ContentType', this.embeddingContentType);
      types.appendChild(newOverride);

      await XmlHelper.writeXmlToArchive(
        this.targetArchive,
        contentTypesPath,
        contentTypesXml,
      );
    }
  }

  private async updateSlideRels(): Promise<void> {
    const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relationships = relXml.getElementsByTagName('Relationship');

    const newTarget = `../embeddings/${path.basename(this.embeddingFile)}`;

    // Update or create the relationship
    let relationshipUpdated = false;
    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].getAttribute('Id') === this.sourceRid) {
        relationships[i].setAttribute('Id', this.createdRid);
        relationships[i].setAttribute('Target', XmlHelper.sanitizeAttr(newTarget));
        relationshipUpdated = true;
        break;
      }
    }

    if (!relationshipUpdated) {
      const newRel = relXml.createElement('Relationship');
      newRel.setAttribute('Id', this.createdRid);
      newRel.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
      );
      newRel.setAttribute('Target', XmlHelper.sanitizeAttr(newTarget));
      relXml.documentElement.appendChild(newRel);
    }

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
  }

  private async updateSlideXml(): Promise<void> {
    // targetSlideFile respects targetType — an OLE object can sit on a
    // slideLayout or slideMaster (e.g. think-cell data objects), where a
    // hardcoded slide path would hit ppt/slides/slide<layoutNumber>.xml.
    const slideXmlPath = this.targetSlideFile;
    const slideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      slideXmlPath,
    );

    const oleObjs = Array.from(slideXml.getElementsByTagName('p:oleObj'));
    oleObjs.forEach((oleObj) => {
      if (oleObj.getAttribute('r:id') === this.sourceRid) {
        oleObj.setAttribute('r:id', this.createdRid);
        const oleObjPr = oleObj.getElementsByTagName('p:oleObjPr')[0];
        if (oleObjPr) {
          const links = Array.from(oleObjPr.getElementsByTagName('a:link'));
          links.forEach((link) => {
            link.setAttribute('r:id', this.createdRid);
          });
        }
      }
    });

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      slideXmlPath,
      slideXml,
    );
  }

  private getContentType(fileExtension: string): string {
    const contentTypes: { [key: string]: string } = {
      '.bin': 'application/vnd.openxmlformats-officedocument.oleObject',
      '.xls': 'application/vnd.ms-excel',
      '.xlsx':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      '.doc': 'application/msword',
      '.docx':
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      '.ppt': 'application/vnd.ms-powerpoint',
      '.pptx':
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    };
    return (
      contentTypes[fileExtension.toLowerCase()] ||
      'application/vnd.openxmlformats-officedocument.oleObject'
    );
  }

  private async removeOleObjectFile(): Promise<void> {
    await this.targetArchive.remove(this.embeddingFile);
  }

  private async removeFromContentTypes(): Promise<void> {
    const contentTypesPath = '[Content_Types].xml';
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      contentTypesPath,
    );

    const types = contentTypesXml.getElementsByTagName('Types')[0];
    const partName = PptPaths.partName(this.embeddingFile);
    const overrideToRemove = Array.from(
      types.getElementsByTagName('Override'),
    ).find((override) => override.getAttribute('PartName') === partName);

    if (overrideToRemove) {
      types.removeChild(overrideToRemove);
      await XmlHelper.writeXmlToArchive(
        this.targetArchive,
        contentTypesPath,
        contentTypesXml,
      );
    }
  }

  private async removeFromSlideRels(): Promise<void> {
    const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relationships = relXml.getElementsByTagName('Relationship');

    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].getAttribute('Id') === this.createdRid) {
        relationships[i].parentNode.removeChild(relationships[i]);
        break;
      }
    }

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    const oleObjectType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject';
    // An embedded Office document (Excel worksheet, Word document) is a
    // "package" relationship, not an "oleObject" one
    const packageType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';

    return XmlRelationshipHelper.getRelationshipItems(
      archive,
      relsPath,
      (element: XmlElement, rels: Target[]) => {
        const type = element.getAttribute('Type');

        // A linked object (TargetMode="External") has no file to copy; its
        // relationship is copied with the slide's rels as it is
        if (element.getAttribute('TargetMode') === 'External') return;
        const isEmbeddedPackage =
          type === packageType &&
          /(^|\/)embeddings\//.test(element.getAttribute('Target'));
        if (type === oleObjectType || isEmbeddedPackage) {
          rels.push({
            rId: element.getAttribute('Id'),
            type: element.getAttribute('Type'),
            file: element.getAttribute('Target'),
            element: element,
          } as Target);
        }
      },
    );
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    oleObjects: Target[],
  ): Promise<void> {
    await this.prepare(targetTemplate, targetSlideNumber, oleObjects);
  }
}
