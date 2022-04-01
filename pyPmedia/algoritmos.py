
from qgis.core import QgsProcessing
from qgis.core import QgsProcessingAlgorithm
from qgis.core import QgsProcessingMultiStepFeedback
from qgis.core import QgsProcessingParameterField
from qgis.core import QgsProcessingParameterVectorLayer
from qgis.core import QgsProcessingParameterFeatureSink
from qgis.core import QgsProcessingParameterBoolean
from qgis.core import QgsProcessingUtils
import processing


class thiessenClip(QgsProcessingAlgorithm):


    def fThiessenClip(self, parameters, context, model_feedback):

        feedback = QgsProcessingMultiStepFeedback(9, model_feedback)
        results = {}
        outputs = {}


        # v.voronoi
        alg_params = {
            '-l': False,
            '-t': False,
            'GRASS_MIN_AREA_PARAMETER': 0.0001,
            'GRASS_OUTPUT_TYPE_PARAMETER': 3,
            'GRASS_REGION_PARAMETER': "%f, %f, %f, %f" % (parameters['box'][0], parameters['box'][1], parameters['box'][2], parameters['box'][3]),
            'GRASS_SNAP_TOLERANCE_PARAMETER': -1,
            'GRASS_VECTOR_DSCO': '',
            'GRASS_VECTOR_EXPORT_NOCAT': False,
            'GRASS_VECTOR_LCO': '',
            'input': parameters['layerGagesSelects'],
            'output': parameters['layerThiessenTemp']
        }
        outputs['voronoi'] = processing.run('grass7:v.voronoi', alg_params, context=None, feedback=None, is_child_algorithm=True)


        alg_params2 = {
            'INPUT': outputs['voronoi']['output'],
            'OVERLAY':parameters['layerWatershed'],
            'OUTPUT':parameters['layerClipTemp']
        }
        outputs['clip'] = processing.run('native:clip', alg_params2, context=None, feedback=None, is_child_algorithm=True)

        results['CLIP'] = outputs['clip']['OUTPUT']
        return results