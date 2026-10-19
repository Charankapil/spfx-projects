import { IExplorerData, ILoadProgress } from '../models/IMetadata';

export interface IMetadataSource {
  load(onProgress: (progress: ILoadProgress) => void): Promise<IExplorerData>;
}
